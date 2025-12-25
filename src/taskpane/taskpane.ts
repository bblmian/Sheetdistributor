/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// 配置列名称常量
const CFG_SN_COL_NAME = "Wiz_Cfg_SN_Col";
const CFG_AMT_COL_NAME = "Wiz_Cfg_Amt_Col";
const CFG_DATA_RANGE_NAME = "Wiz_Cfg_Data_Range";
const CFG_HEADER_ROW_NAME = "Wiz_Cfg_Header_Row";

// 主报表配置接口
interface MainReportConfig {
  dataRange: string; // 数据区域地址，如 "Sheet1!A1:E100"
  headerRow: number; // 标题行号（1-based）
  snColumn: number; // S/N 列索引（0-based）
  amtColumn: number; // 金额列索引（0-based）
  sheetName: string; // 工作表名称
}

// 将列索引转换为 Excel 列名（支持 A-Z 和 AA-ZZ）
function getColumnName(columnIndex: number): string {
  let result = "";
  let index = columnIndex;
  while (index >= 0) {
    result = String.fromCharCode(65 + (index % 26)) + result;
    index = Math.floor(index / 26) - 1;
  }
  return result;
}

// 从 Excel 列名转换为列索引（例如 "A" -> 0, "Z" -> 25, "AA" -> 26）
function getColumnIndexFromName(columnName: string): number {
  let index = 0;
  columnName = columnName.toUpperCase();
  for (let i = 0; i < columnName.length; i++) {
    index = index * 26 + (columnName.charCodeAt(i) - 64);
  }
  return index - 1;
}

// 从 NamedItem 公式中解析列名（例如 "=Sheet1!$A:$A" -> "A"）
function parseColumnFromFormula(formula: string): string | null {
  // 匹配格式：=SheetName!$Column:$Column 或 =SheetName!Column:Column
  const match = formula.match(/![\$]?([A-Z]+)[\$]?:[\$]?[A-Z]+/i);
  if (match && match[1]) {
    return match[1].toUpperCase();
  }
  return null;
}

// 报表信息接口
interface ReportInfo {
  sheetName: string;
  createTime: string;
  rowCount: number;
  totalAmount: number;
}

// 列筛选条件接口
interface ColumnFilterSetting {
  columnIndex: number;
  columnName: string; // 列名（如 "A", "B"）
  headerText: string; // 标题行文本
  filterType: string; // 筛选类型
  filterValues: string[]; // 筛选值列表
  isFiltered: boolean; // 是否有筛选
}

// 筛选条件信息接口
interface FilterCondition {
  id: string;
  name: string;
  createTime: string;
  sheetName: string;
  filterSettings: ColumnFilterSetting[]; // 存储每列的筛选设置
  config: MainReportConfig; // 主报表配置
}

// 存储报表列表（使用内存存储，因为 Excel Add-in 的 localStorage 可能受限）
let reportsList: ReportInfo[] = [];

// 存储筛选条件列表
let filterConditionsList: FilterCondition[] = [];

// 当前主报表配置
let currentMainReportConfig: MainReportConfig | null = null;

// 筛选面板字段数据接口
interface FilterFieldData {
  columnIndex: number;
  headerText: string;
  allValues: string[]; // 该列所有唯一值
  selectedValues: Set<string>; // 用户选中的值
}

// 当前筛选面板的字段数据
let filterPanelFields: FilterFieldData[] = [];

// 筛选统计数据接口
interface FilterStatistics {
  filteredRowCount: number;   // 筛选出的行数
  totalAmount: number;        // 合计金额
  isValid: boolean;           // 是否有效（是否已计算）
}

// 当前筛选的统计数据
let currentFilterStatistics: FilterStatistics = {
  filteredRowCount: 0,
  totalAmount: 0,
  isValid: false
};

// 显示消息到消息区域
function showMessage(message: string, isError: boolean = false): void {
  const messageArea = document.getElementById("message-area");
  if (messageArea) {
    messageArea.textContent = message;
    messageArea.className = isError ? "message-area error" : "message-area success";
  }
  console.log(message);
}

// 显示进度条
function showProgress(title: string, message: string): void {
  const overlay = document.getElementById("progress-overlay");
  const titleEl = document.getElementById("progress-title");
  const messageEl = document.getElementById("progress-message");
  const barEl = document.getElementById("progress-bar");
  const detailEl = document.getElementById("progress-detail");
  
  if (overlay && titleEl && messageEl && barEl && detailEl) {
    titleEl.textContent = title;
    messageEl.textContent = message;
    barEl.style.width = "0%";
    detailEl.textContent = "";
    overlay.classList.add("show");
  }
}

// 更新进度条
function updateProgress(current: number, total: number, detail?: string): void {
  const barEl = document.getElementById("progress-bar");
  const detailEl = document.getElementById("progress-detail");
  
  if (barEl && detailEl) {
    const percent = total > 0 ? Math.round((current / total) * 100) : 0;
    barEl.style.width = `${percent}%`;
    detailEl.textContent = detail || `${current} / ${total}`;
  }
}

// 隐藏进度条
function hideProgress(): void {
  const overlay = document.getElementById("progress-overlay");
  if (overlay) {
    overlay.classList.remove("show");
  }
}

// 追加调试日志到界面
function appendDebugLog(message: string): void {
  const debugLog = document.getElementById("debug-log");
  if (debugLog) {
    const logEntry = document.createElement("div");
    logEntry.textContent = message;
    logEntry.style.marginBottom = "2px";
    debugLog.appendChild(logEntry);
    // 自动滚动到底部
    debugLog.scrollTop = debugLog.scrollHeight;
  }
  console.log(message);
}

// 生成筛选条件的文本格式
function generateFilterText(): string {
  const filterItems: string[] = [];
  
  // 首先添加统计数据
  if (currentFilterStatistics.isValid) {
    filterItems.push(`筛选出的总条数：${currentFilterStatistics.filteredRowCount}`);
    filterItems.push(`对应合计金额：${currentFilterStatistics.totalAmount.toFixed(2)}`);
    filterItems.push(""); // 空行分隔
  }
  
  // 添加筛选条件
  let hasFilter = false;
  for (const field of filterPanelFields) {
    // 只显示有筛选的字段（选中的值少于全部值，且至少选中一个）
    if (field.selectedValues.size < field.allValues.length && field.selectedValues.size > 0) {
      const valuesArray = Array.from(field.selectedValues);
      const valuesText = valuesArray.length > 5 
        ? valuesArray.slice(0, 5).join("、") + `...等${valuesArray.length}项`
        : valuesArray.join("、");
      filterItems.push(`【${field.headerText}】：${valuesText}`);
      hasFilter = true;
    }
  }
  
  if (!hasFilter && !currentFilterStatistics.isValid) {
    return "无筛选条件";
  }
  
  return filterItems.join("\n");
}

// 更新筛选条件文本区域
function updateFilterTextDisplay(autoHide: boolean = false): void {
  const container = document.getElementById("filter-text-container");
  const content = document.getElementById("filter-text-content");
  
  if (!container || !content) return;
  
  const filterText = generateFilterText();
  
  if (filterText === "无筛选条件" && autoHide) {
    container.style.display = "none";
    return;
  }
  
  // 显示容器
  container.style.display = "block";
  
  // 生成带样式的 HTML 内容
  const htmlContent = filterText.split("\n").map(line => {
    // 空行处理
    if (line.trim() === "") {
      return "<hr class='filter-text-divider'>";
    }
    
    // 统计数据行：筛选出的总条数 或 对应合计金额
    if (line.startsWith("筛选出的总条数：")) {
      const value = line.replace("筛选出的总条数：", "");
      return `<span class="stat-label">筛选出的总条数：</span><span class="stat-value">${value}</span>`;
    }
    if (line.startsWith("对应合计金额：")) {
      const value = line.replace("对应合计金额：", "");
      return `<span class="stat-label">对应合计金额：</span><span class="stat-value">${value}</span>`;
    }
    
    // 解析 【字段名】：值 格式
    const match = line.match(/【(.+?)】：(.+)/);
    if (match) {
      return `<span class="field-name">【${match[1]}】</span>：<span class="field-value">${match[2]}</span>`;
    }
    return line;
  }).join("<br>");
  
  content.innerHTML = htmlContent;
}

// 复制筛选条件到剪贴板
async function copyFilterTextToClipboard(): Promise<boolean> {
  const filterText = generateFilterText();
  
  if (filterText === "无筛选条件") {
    return false;
  }
  
  try {
    await navigator.clipboard.writeText(filterText);
    showCopySuccessHint();
    return true;
  } catch (error) {
    console.error("复制到剪贴板失败:", error);
    // 备用方案：使用传统的复制方式
    try {
      const textArea = document.createElement("textarea");
      textArea.value = filterText;
      textArea.style.position = "fixed";
      textArea.style.left = "-9999px";
      document.body.appendChild(textArea);
      textArea.select();
      document.execCommand("copy");
      document.body.removeChild(textArea);
      showCopySuccessHint();
      return true;
    } catch (fallbackError) {
      console.error("备用复制方式也失败:", fallbackError);
      return false;
    }
  }
}

// 显示复制成功提示
function showCopySuccessHint(): void {
  const hint = document.getElementById("copy-success-hint");
  if (hint) {
    hint.classList.add("show");
    setTimeout(() => {
      hint.classList.remove("show");
    }, 2000);
  }
}

// 绑定筛选条件文本区域的点击事件
function bindFilterTextEvents(): void {
  const content = document.getElementById("filter-text-content");
  if (content) {
    content.addEventListener("click", async () => {
      await copyFilterTextToClipboard();
    });
  }
}

// 清空调试日志
function clearDebugLog(): void {
  const debugLog = document.getElementById("debug-log");
  if (debugLog) {
    debugLog.innerHTML = "";
  }
}

// 清除货币符号和其他干扰字符，转换为数字
function cleanAmount(value: any): number {
  if (value === null || value === undefined || value === "") {
    return 0;
  }
  
  // 如果是数字，直接返回
  if (typeof value === "number") {
    return isNaN(value) ? 0 : value;
  }
  
  // 转换为字符串并清理
  let str = String(value).trim();
  
  // 如果字符串为空，返回0
  if (str === "") {
    return 0;
  }
  
  // 移除 RMB、CNY 等货币单位（不区分大小写）
  str = str.replace(/RMB|CNY|USD|EUR|GBP/gi, "");
  
  // 移除常见的货币符号（包括中文和英文的 ¥、$、€、£、￥）
  str = str.replace(/[¥$€£￥]/g, "");
  
  // 移除中文和英文的千位分隔符（逗号）
  str = str.replace(/[,，]/g, "");
  
  // 移除所有空格
  str = str.replace(/\s+/g, "");
  
  // 处理负数（保留开头的负号）
  const isNegative = str.startsWith("-");
  if (isNegative) {
    str = str.substring(1);
  }
  
  // 只保留数字和小数点
  str = str.replace(/[^\d.]/g, "");
  
  // 处理多个小数点的情况（只保留第一个）
  const dotIndex = str.indexOf(".");
  if (dotIndex !== -1) {
    str = str.substring(0, dotIndex + 1) + str.substring(dotIndex + 1).replace(/\./g, "");
  }
  
  // 转换为数字
  const num = parseFloat(str);
  const result = isNaN(num) ? 0 : (isNegative ? -num : num);
  
  return result;
}

// 设置主报表数据区域
async function setDataRange(): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      
      // 获取当前选中的范围
      const selection = context.workbook.getSelectedRange();
      if (!selection) {
        showMessage("请先选择一个数据区域", true);
        return;
      }
      
      selection.load("address, rowIndex, columnIndex, rowCount, columnCount");
      await context.sync();
      
      // 获取地址
      let address = selection.address;
      
      // 如果地址包含工作表名称，提取出来
      let sheetNameInAddress = "";
      if (address.includes("!")) {
        const parts = address.split("!");
        if (parts.length > 1) {
          sheetNameInAddress = parts[0].replace(/^'|'$/g, "");
          address = parts[1];
        }
      }
      
      // 确保地址使用绝对引用格式（添加 $ 符号）
      // 如果地址已经是绝对引用，保持不变；否则转换为绝对引用
      if (!address.includes("$")) {
        // 手动构建绝对引用地址
        const startRow = selection.rowIndex + 1;
        const startCol = selection.columnIndex;
        const endRow = startRow + selection.rowCount - 1;
        const endCol = startCol + selection.columnCount - 1;
        
        const startColName = getColumnName(startCol);
        const endColName = getColumnName(endCol);
        
        address = `$${startColName}$${startRow}:$${endColName}$${endRow}`;
      }
      
      // 使用当前工作表名称（而不是地址中的工作表名称）
      const currentSheetName = sheet.name;
      // 如果工作表名称包含空格或特殊字符，需要用单引号包裹
      const sheetNameEscaped = currentSheetName.includes(" ") || currentSheetName.includes("-") || currentSheetName.includes("'")
        ? `'${currentSheetName.replace(/'/g, "''")}'` 
        : currentSheetName;
      
      const fullAddress = `${sheetNameEscaped}!${address}`;
      
      // 保存到 NamedItem
      const namedItems = context.workbook.names;
      
      // 先尝试删除已存在的（如果存在）
      try {
        const existingItem = namedItems.getItem(CFG_DATA_RANGE_NAME);
        existingItem.delete();
        await context.sync();
      } catch (error) {
        // 如果不存在，忽略错误
      }
      
      // 创建 NamedItem，使用公式字符串（必须以 = 开头）
      const formula = `=${fullAddress}`;
      
      // 添加调试信息
      appendDebugLog(`创建 NamedItem: ${CFG_DATA_RANGE_NAME}`);
      appendDebugLog(`公式: ${formula}`);
      appendDebugLog(`地址: ${fullAddress}`);
      
      try {
        namedItems.add(CFG_DATA_RANGE_NAME, formula);
        await context.sync();
      } catch (addError) {
        console.error("创建 NamedItem 失败:", addError);
        appendDebugLog(`错误: ${addError.message}`);
        throw addError;
      }
      
      // 更新配置
      if (!currentMainReportConfig) {
        currentMainReportConfig = {
          dataRange: fullAddress,
          headerRow: 0,
          snColumn: 0,
          amtColumn: 0,
          sheetName: currentSheetName
        };
      } else {
        currentMainReportConfig.dataRange = fullAddress;
        currentMainReportConfig.sheetName = currentSheetName;
      }
      
      updateMainReportConfigDisplay();
      showMessage(`已设置数据区域: ${address}`);
    });
  } catch (error) {
    console.error("设置数据区域时出错:", error);
    showMessage(`设置数据区域失败: ${error.message}`, true);
  }
}

// 设置标题行
async function setHeaderRow(): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      
      // 获取当前选中的范围
      const selection = context.workbook.getSelectedRange();
      selection.load("rowIndex");
      
      await context.sync();
      
      const rowIndex = selection.rowIndex + 1; // 转换为 1-based
      
      // 保存到 NamedItem
      const namedItems = context.workbook.names;
      try {
        const existingItem = namedItems.getItem(CFG_HEADER_ROW_NAME);
        existingItem.delete();
        await context.sync();
      } catch (error) {
        // 如果不存在，忽略错误
      }
      
      const rowAddress = `${sheet.name}!${rowIndex}:${rowIndex}`;
      namedItems.add(CFG_HEADER_ROW_NAME, `=${rowAddress}`);
      await context.sync();
      
      // 更新配置
      if (!currentMainReportConfig) {
        currentMainReportConfig = {
          dataRange: "",
          headerRow: rowIndex,
          snColumn: 0,
          amtColumn: 0,
          sheetName: sheet.name
        };
      } else {
        currentMainReportConfig.headerRow = rowIndex;
        currentMainReportConfig.sheetName = sheet.name;
      }
      
      updateMainReportConfigDisplay();
      showMessage(`已设置标题行: 第 ${rowIndex} 行`);
    });
  } catch (error) {
    console.error("设置标题行时出错:", error);
    showMessage(`设置标题行失败: ${error.message}`, true);
  }
}

// 设置列配置（SN列和金额列）
async function setColumnConfig(columnName: string, displayName: string): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      
      // 获取当前选中的范围
      const selection = context.workbook.getSelectedRange();
      selection.load("address, columnIndex");
      
      await context.sync();
      
      // 验证选中范围是否有效
      if (!selection.address) {
        showMessage(`请先在 Excel 中选择${displayName}所在的列中的任意单元格`, true);
        return;
      }
      
      // 直接使用 selection 的 columnIndex
      const columnIndex = selection.columnIndex;
      const columnNameStr = getColumnName(columnIndex);
      
      // 创建或更新 NamedItem
      const namedItems = context.workbook.names;
      
      // 先尝试删除已存在的（如果存在）
      try {
        const existingItem = namedItems.getItemOrNullObject(columnName);
        await context.sync();
        if (!existingItem.isNullObject) {
          existingItem.delete();
          await context.sync();
        }
      } catch (error) {
        // 如果删除失败，忽略错误
        console.log("删除已有 NamedItem 时出错（可能不存在）:", error);
      }
      
      // 创建新的 NamedItem，引用到选中的单元格所在列的第一个单元格
      // 使用单元格引用而不是整列引用，避免某些 Excel 版本的兼容问题
      const cellAddress = `='${sheet.name}'!$${columnNameStr}$1`;
      namedItems.add(columnName, cellAddress);
      
      await context.sync();
      
      // 更新配置
      if (!currentMainReportConfig) {
        currentMainReportConfig = {
          dataRange: "",
          headerRow: 0,
          snColumn: columnName === CFG_SN_COL_NAME ? columnIndex : 0,
          amtColumn: columnName === CFG_AMT_COL_NAME ? columnIndex : 0,
          sheetName: sheet.name
        };
      } else {
        if (columnName === CFG_SN_COL_NAME) {
          currentMainReportConfig.snColumn = columnIndex;
        } else if (columnName === CFG_AMT_COL_NAME) {
          currentMainReportConfig.amtColumn = columnIndex;
        }
        currentMainReportConfig.sheetName = sheet.name;
      }
      
      showMessage(`成功设置${displayName}为第 ${columnIndex + 1} 列 (${columnNameStr})`);
      
      // 更新配置显示
      updateMainReportConfigDisplay();
    });
  } catch (error) {
    console.error("设置列配置时出错:", error);
    showMessage(`设置${displayName}失败: ${error.message}`, true);
  }
}

// 设置SN列（带合并单元格检查）
// 合并单元格信息接口
interface MergeCellInfo {
  address: string;      // 合并区域地址，如 "A5:A6"
  startRow: number;     // 起始行（0-based）
  endRow: number;       // 结束行（0-based）
  columnIndex: number;  // 列索引
}

async function setSnColumnWithMergeCheck(): Promise<void> {
  try {
    // 立即显示进度条
    showProgress("检查合并单元格", "正在读取选中列数据...");
    
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
      sheet.load("name");
      
      // 获取当前选中的范围
      const selection = context.workbook.getSelectedRange();
      selection.load("address, columnIndex");
      
      // 获取该列的使用范围
      const usedRange = sheet.getUsedRange();
      usedRange.load("rowCount, rowIndex, columnCount, address");
      
      await context.sync();
      
      const columnIndex = selection.columnIndex;
      const columnNameStr = getColumnName(columnIndex);
      const startRow = usedRange.rowIndex;
      const rowCount = usedRange.rowCount;
      const totalColumns = usedRange.columnCount;
      
      console.log(`开始检查列 ${columnNameStr} 的合并单元格，范围: 行${startRow+1}到${startRow+rowCount}`);
      
      // 更新进度条
      updateProgress(10, 100, `正在扫描第 ${columnIndex + 1} 列 (${columnNameStr})，共 ${rowCount} 行...`);
      
      // 读取整列的值 - 一次性加载
      const columnRange = sheet.getRangeByIndexes(startRow, columnIndex, rowCount, 1);
      columnRange.load("values");
      await context.sync();
      
      const values = columnRange.values;
      console.log(`读取到 ${values.length} 个单元格值`);
      
      // 更新进度条
      updateProgress(30, 100, `数据读取完成，正在检测合并单元格...`);
      
      // ===== 使用 Excel API 直接检测合并单元格 =====
      const mergedCellInfos: MergeCellInfo[] = [];
      
      // 方法1：使用 getMergedAreasOrNullObject 直接检测整列的合并区域
      updateProgress(40, 100, "正在使用 API 检测合并单元格...");
      
      try {
        const mergedAreas = columnRange.getMergedAreasOrNullObject();
        mergedAreas.load("isNullObject, areaCount");
        await context.sync();
        
        if (!mergedAreas.isNullObject && mergedAreas.areaCount > 0) {
          // 加载每个合并区域的信息
          mergedAreas.load("areas");
          await context.sync();
          
          const areas = mergedAreas.areas;
          areas.load("items");
          await context.sync();
          
          for (let j = 0; j < areas.items.length; j++) {
            const area = areas.items[j];
            area.load("rowCount, address, rowIndex, columnIndex");
          }
          await context.sync();
          
          for (let j = 0; j < areas.items.length; j++) {
            const area = areas.items[j];
            if (area.rowCount > 1 && area.columnIndex === columnIndex) {
              mergedCellInfos.push({
                address: area.address,
                startRow: area.rowIndex,
                endRow: area.rowIndex + area.rowCount - 1,
                columnIndex: columnIndex
              });
              console.log(`API 检测到合并区域: ${area.address}, 跨${area.rowCount}行`);
            }
          }
        }
      } catch (apiError) {
        console.log("API 检测方法失败，使用备用方法:", apiError);
      }
      
      // 方法2（备用）：如果 API 方法没有检测到，使用值模式检测
      if (mergedCellInfos.length === 0) {
        updateProgress(50, 100, "使用值模式检测合并单元格...");
        
        // 合并单元格的特征：主单元格有值或为空，其下方的从属单元格值为 null 或空字符串
        let i = 0;
        
        while (i < rowCount) {
          const currentValue = values[i][0];
          
          // 检查后续是否有连续的 null 或空值
          let consecutiveEmpty = 0;
          for (let k = i + 1; k < rowCount; k++) {
            const nextValue = values[k][0];
            // null 或 undefined 或空字符串都视为可能的合并从属单元格
            if (nextValue === null || nextValue === undefined || nextValue === "") {
              consecutiveEmpty++;
            } else {
              break;
            }
          }
          
          // 只有当主单元格有值且后续有连续空值时，才认为是合并单元格
          if (consecutiveEmpty > 0 && currentValue !== null && currentValue !== undefined && currentValue !== "") {
            const cellRowIndex = startRow + i;
            const mergeRowCount = consecutiveEmpty + 1;
            const endRowIndex = cellRowIndex + mergeRowCount - 1;
            const address = `${columnNameStr}${cellRowIndex + 1}:${columnNameStr}${endRowIndex + 1}`;
            
            console.log(`值模式检测到可能的合并区域: ${address}, 跨${mergeRowCount}行`);
            
            mergedCellInfos.push({
              address: address,
              startRow: cellRowIndex,
              endRow: endRowIndex,
              columnIndex: columnIndex
            });
            
            i += mergeRowCount;
          } else {
            i++;
          }
        }
      }
      
      console.log(`总共检测到 ${mergedCellInfos.length} 个合并单元格区域`);
      
      // 更新进度条
      updateProgress(80, 100, `检测完成，发现 ${mergedCellInfos.length} 个合并区域`);
      
      // 隐藏进度条
      hideProgress();
      
      // 如果存在合并单元格，提示用户
      if (mergedCellInfos.length > 0) {
        showMessage(`检测到 ${mergedCellInfos.length} 个合并单元格区域`, false);
        
        // 询问是否标注合并单元格
        const confirmHighlight = await showConfirmDialog(
          `在第 ${columnIndex + 1} 列 (${columnNameStr}) 中检测到 ${mergedCellInfos.length} 个合并单元格区域。\n\n是否将这些合并单元格用黄色背景标注？`
        );
        
        if (confirmHighlight) {
          // 用黄色标注合并单元格
          for (const info of mergedCellInfos) {
            const mergeRange = sheet.getRange(info.address);
            mergeRange.format.fill.color = "#FFFF00";
          }
          await context.sync();
          showMessage(`已标注 ${mergedCellInfos.length} 个合并单元格区域`);
        }
        
        // 询问是否合并行数据
        const confirmMerge = await showConfirmDialog(
          `是否需要拆分合并单元格并合并行数据？\n\n操作说明：\n• 取消单元格合并\n• 相同内容保留其一\n• 不同内容换行拼接到第一行\n• 删除多余行\n• 用浅灰色背景标注处理过的行`
        );
        
        if (confirmMerge) {
          // 显示进度条
          showProgress("处理合并单元格", `正在处理 ${mergedCellInfos.length} 个合并区域，请耐心等待...`);
          
          // 从后往前处理，避免删除行后索引错乱
          const sortedInfos = mergedCellInfos.sort((a, b) => b.startRow - a.startRow);
          let deletedRows = 0;
          const totalInfos = sortedInfos.length;
          
          // ===== 优化：批量读取所有需要的数据 =====
          // 先计算需要读取的最大范围
          let minRow = Infinity;
          let maxRow = 0;
          for (const info of sortedInfos) {
            minRow = Math.min(minRow, info.startRow);
            maxRow = Math.max(maxRow, info.endRow);
          }
          
          // 一次性读取所有相关行的数据
          const allDataRange = sheet.getRangeByIndexes(minRow, 0, maxRow - minRow + 1, totalColumns);
          allDataRange.load("values");
          await context.sync();
          const allData = allDataRange.values;
          
          updateProgress(0, totalInfos, "数据读取完成，开始处理...");
          
          // 收集要处理的结果
          interface MergeResult {
            firstRowIndex: number;
            mergedValues: (string | number | boolean | null)[];
            rowsToDelete: number;
          }
          const mergeResults: MergeResult[] = [];
          
          // 在内存中处理合并逻辑
          for (let idx = 0; idx < sortedInfos.length; idx++) {
            const info = sortedInfos[idx];
            const rowsToMerge = info.endRow - info.startRow + 1;
            if (rowsToMerge < 2) continue;
            
            const firstRowIndex = info.startRow;
            const dataOffset = firstRowIndex - minRow;
            
            // 复制第一行数据
            const mergedValues = allData[dataOffset].slice();
            
            // 合并其他行的数据
            for (let rowOffset = 1; rowOffset < rowsToMerge; rowOffset++) {
              const currentDataOffset = dataOffset + rowOffset;
              if (currentDataOffset >= allData.length) break;
              
              const currentRowValues = allData[currentDataOffset];
              
              for (let col = 0; col < totalColumns; col++) {
                const firstVal = String(mergedValues[col] || "").trim();
                const currentVal = String(currentRowValues[col] || "").trim();
                
                if (currentVal && currentVal !== firstVal) {
                  if (firstVal) {
                    mergedValues[col] = firstVal + "\n" + currentVal;
                  } else {
                    mergedValues[col] = currentVal;
                  }
                }
              }
            }
            
            mergeResults.push({
              firstRowIndex,
              mergedValues,
              rowsToDelete: rowsToMerge - 1
            });
          }
          
          // ===== 批量执行：取消合并、写入数据、删除行 =====
          // 分批处理，每批处理一定数量，避免单次操作太多
          const BATCH_SIZE = 20;
          
          for (let batchStart = 0; batchStart < mergeResults.length; batchStart += BATCH_SIZE) {
            const batchEnd = Math.min(batchStart + BATCH_SIZE, mergeResults.length);
            const batchResults = mergeResults.slice(batchStart, batchEnd);
            
            updateProgress(batchStart, mergeResults.length, `批量处理 ${batchStart + 1}-${batchEnd}/${mergeResults.length}`);
            
            // 1. 批量取消合并
            for (const result of batchResults) {
              const info = sortedInfos.find(i => i.startRow === result.firstRowIndex);
              if (info) {
                const mergeRange = sheet.getRange(info.address);
                mergeRange.unmerge();
              }
            }
            await context.sync();
            
            // 2. 批量写入合并后的数据
            for (const result of batchResults) {
              const firstRowRange = sheet.getRangeByIndexes(result.firstRowIndex, 0, 1, totalColumns);
              firstRowRange.values = [result.mergedValues];
              firstRowRange.format.wrapText = true;
              firstRowRange.format.fill.color = "#F5F5F5";
            }
            await context.sync();
            
            // 3. 批量删除行（从后往前）
            for (const result of batchResults) {
              for (let i = result.rowsToDelete; i > 0; i--) {
                const rowToDelete = sheet.getRangeByIndexes(result.firstRowIndex + i, 0, 1, totalColumns);
                rowToDelete.delete(Excel.DeleteShiftDirection.up);
                deletedRows++;
              }
            }
            await context.sync();
          }
          
          // 隐藏进度条
          hideProgress();
          
          showMessage(`已拆分并合并 ${mergedCellInfos.length} 个区域的行数据，删除了 ${deletedRows} 行`);
        }
      } else {
        // 没有检测到合并单元格，给出提示
        showMessage(`第 ${columnIndex + 1} 列 (${columnNameStr}) 未检测到合并单元格，共扫描 ${rowCount} 行`, false);
      }
      
      // 继续设置SN列
      const namedItems = context.workbook.names;
      
      try {
        const existingItem = namedItems.getItem(CFG_SN_COL_NAME);
        existingItem.delete();
        await context.sync();
      } catch (error) {
        // 如果不存在，忽略错误
      }
      
      const columnAddress = `='${sheet.name}'!$${columnNameStr}:$${columnNameStr}`;
      namedItems.add(CFG_SN_COL_NAME, columnAddress);
      
      await context.sync();
      
      if (!currentMainReportConfig) {
        currentMainReportConfig = {
          dataRange: "",
          headerRow: 0,
          snColumn: columnIndex,
          amtColumn: 0,
          sheetName: sheet.name
        };
      } else {
        currentMainReportConfig.snColumn = columnIndex;
        currentMainReportConfig.sheetName = sheet.name;
      }
      
      const mergeInfo = mergedCellInfos.length > 0 ? ` (已处理 ${mergedCellInfos.length} 个合并区域)` : "";
      showMessage(`成功设置 S/N 列为第 ${columnIndex + 1} 列 (${columnNameStr})${mergeInfo}`);
      
      updateMainReportConfigDisplay();
    });
  } catch (error) {
    console.error("设置SN列时出错:", error);
    hideProgress(); // 确保隐藏进度条
    showMessage(`设置 S/N 列失败: ${error.message}`, true);
  }
}

// 更新主报表配置显示
function updateMainReportConfigDisplay(): void {
  // 更新数据区域显示
  const dataRangeEl = document.getElementById("data-range-info");
  if (dataRangeEl) {
    if (currentMainReportConfig && currentMainReportConfig.dataRange) {
      dataRangeEl.textContent = currentMainReportConfig.dataRange;
      dataRangeEl.className = "step-value clickable";
      dataRangeEl.title = "点击跳转到主报表";
    } else {
      dataRangeEl.textContent = "未配置";
      dataRangeEl.className = "step-value not-configured";
      dataRangeEl.title = "";
    }
  }
  
  // 更新标题行显示
  const headerRowEl = document.getElementById("header-row-info");
  if (headerRowEl) {
    if (currentMainReportConfig && currentMainReportConfig.headerRow > 0) {
      headerRowEl.textContent = `第 ${currentMainReportConfig.headerRow} 行`;
      headerRowEl.className = "step-value";
    } else {
      headerRowEl.textContent = "未配置";
      headerRowEl.className = "step-value not-configured";
    }
  }
  
  // 更新 S/N 列显示
  const snColEl = document.getElementById("sn-col-info");
  if (snColEl) {
    if (currentMainReportConfig && currentMainReportConfig.snColumn >= 0) {
      const colName = getColumnName(currentMainReportConfig.snColumn);
      snColEl.textContent = `第 ${currentMainReportConfig.snColumn + 1} 列 (${colName})`;
      snColEl.className = "step-value";
    } else {
      snColEl.textContent = "未配置";
      snColEl.className = "step-value not-configured";
    }
  }
  
  // 更新金额列显示
  const amtColEl = document.getElementById("amt-col-info");
  if (amtColEl) {
    if (currentMainReportConfig && currentMainReportConfig.amtColumn >= 0) {
      const colName = getColumnName(currentMainReportConfig.amtColumn);
      amtColEl.textContent = `第 ${currentMainReportConfig.amtColumn + 1} 列 (${colName})`;
      amtColEl.className = "step-value";
    } else {
      amtColEl.textContent = "未配置";
      amtColEl.className = "step-value not-configured";
    }
  }
  
  // 更新启用筛选按钮状态
  const enableFilterBtn = document.getElementById("btn-enable-filter") as HTMLButtonElement;
  if (enableFilterBtn) {
    const isConfigComplete = currentMainReportConfig && 
      currentMainReportConfig.dataRange && 
      currentMainReportConfig.headerRow > 0 &&
      currentMainReportConfig.snColumn >= 0 &&
      currentMainReportConfig.amtColumn >= 0;
    enableFilterBtn.disabled = !isConfigComplete;
  }
}

// 更新列配置显示（保留用于兼容）
async function updateColumnConfigDisplay(): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      
      // 获取 S/N 列信息
      const snInfoEl = document.getElementById("sn-col-info");
      try {
        const snNamedItem = workbook.names.getItem(CFG_SN_COL_NAME);
        snNamedItem.load("formula");
        const snRange = snNamedItem.getRange();
        snRange.load("columnIndex");
        await context.sync();
        
        const snColName = getColumnName(snRange.columnIndex);
        if (snInfoEl) {
          snInfoEl.textContent = `第 ${snRange.columnIndex + 1} 列 (${snColName})`;
          snInfoEl.className = "config-value";
        }
      } catch (error) {
        if (snInfoEl) {
          snInfoEl.textContent = "未配置";
          snInfoEl.className = "config-value not-configured";
        }
      }
      
      // 获取金额列信息
      const amtInfoEl = document.getElementById("amt-col-info");
      try {
        const amtNamedItem = workbook.names.getItem(CFG_AMT_COL_NAME);
        amtNamedItem.load("formula");
        const amtRange = amtNamedItem.getRange();
        amtRange.load("columnIndex");
        await context.sync();
        
        const amtColName = getColumnName(amtRange.columnIndex);
        if (amtInfoEl) {
          amtInfoEl.textContent = `第 ${amtRange.columnIndex + 1} 列 (${amtColName})`;
          amtInfoEl.className = "config-value";
        }
      } catch (error) {
        if (amtInfoEl) {
          amtInfoEl.textContent = "未配置";
          amtInfoEl.className = "config-value not-configured";
        }
      }
    });
  } catch (error) {
    console.error("更新列配置显示时出错:", error);
  }
}

// 更新报表列表显示
function updateReportsTable(): void {
  const tbody = document.getElementById("reports-table-body");
  if (!tbody) return;
  
  let html = "";
  
  // 首先添加主报表行（如果已配置）
  if (currentMainReportConfig && currentMainReportConfig.sheetName) {
    html += `
      <tr class="main-report-row">
        <td>
          <span class="main-report-link" title="点击跳转到主报表">
            <i class="ms-Icon ms-Icon--Home" style="margin-right: 4px; font-size: 10px;"></i>${currentMainReportConfig.sheetName}
          </span>
        </td>
        <td><span class="main-report-badge">主报表</span></td>
        <td></td>
      </tr>
    `;
  }
  
  if (reportsList.length === 0 && !currentMainReportConfig?.sheetName) {
    tbody.innerHTML = '<tr><td colspan="3" class="no-reports">暂无报表</td></tr>';
    return;
  }
  
  html += reportsList.map((report, index) => {
    const createTime = new Date(report.createTime).toLocaleString("zh-CN", { 
      month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit" 
    });
    return `
      <tr>
        <td>
          <span class="report-name-display" data-index="${index}" data-sheet-name="${report.sheetName}" title="点击跳转，双击修改名称">${report.sheetName}</span>
          <div class="report-name-edit-container" style="display: none;">
            <input type="text" class="report-name-input" data-index="${index}" value="${report.sheetName}">
            <i class="ms-Icon ms-Icon--ReturnKey edit-hint-icon" title="按回车确认"></i>
          </div>
        </td>
        <td>${createTime}</td>
        <td>
          <button class="icon-button-small danger" data-index="${index}" data-sheet-name="${report.sheetName}" title="删除报表">
            <i class="ms-Icon ms-Icon--Delete"></i>
          </button>
        </td>
      </tr>
    `;
  }).join("");
  
  tbody.innerHTML = html;
  
  // 绑定主报表点击事件
  tbody.querySelectorAll(".main-report-link").forEach(link => {
    link.addEventListener("click", () => {
      goToMainReport();
    });
  });
  
  // 绑定报表名称单击跳转事件
  tbody.querySelectorAll(".report-name-display").forEach(span => {
    span.addEventListener("click", (e) => {
      const sheetName = (e.target as HTMLElement).getAttribute("data-sheet-name");
      if (sheetName) {
        navigateToSheet(sheetName);
      }
    });
    
    // 绑定双击编辑事件
    span.addEventListener("dblclick", (e) => {
      const target = e.target as HTMLElement;
      const index = parseInt(target.getAttribute("data-index") || "0");
      startEditReportName(index);
    });
  });
  
  // 绑定输入框事件
  tbody.querySelectorAll(".report-name-input").forEach(input => {
    // 回车确认
    input.addEventListener("keydown", (e) => {
      const keyEvent = e as KeyboardEvent;
      if (keyEvent.key === "Enter") {
        const target = e.target as HTMLInputElement;
        const index = parseInt(target.getAttribute("data-index") || "0");
        confirmEditReportName(index, target.value);
      } else if (keyEvent.key === "Escape") {
        const target = e.target as HTMLInputElement;
        const index = parseInt(target.getAttribute("data-index") || "0");
        cancelEditReportName(index);
      }
    });
    
    // 失去焦点取消编辑
    input.addEventListener("blur", (e) => {
      const target = e.target as HTMLInputElement;
      const index = parseInt(target.getAttribute("data-index") || "0");
      // 延迟取消，避免与回车冲突
      setTimeout(() => {
        cancelEditReportName(index);
      }, 200);
    });
  });
  
  // 绑定删除事件
  tbody.querySelectorAll(".icon-button-small.danger").forEach(button => {
    button.addEventListener("click", (e) => {
      // 使用 currentTarget 而不是 target，因为点击图标时 target 是图标
      const btn = e.currentTarget as HTMLElement;
      const index = parseInt(btn.getAttribute("data-index") || "0");
      const sheetName = btn.getAttribute("data-sheet-name");
      if (sheetName) {
        deleteReport(index, sheetName);
      }
    });
  });
}

// 开始编辑报表名称
function startEditReportName(index: number): void {
  const tbody = document.getElementById("reports-table-body");
  if (!tbody) return;
  
  // 隐藏显示区域，显示编辑区域
  const displaySpan = tbody.querySelector(`.report-name-display[data-index="${index}"]`) as HTMLElement;
  const editContainer = displaySpan?.parentElement?.querySelector(".report-name-edit-container") as HTMLElement;
  const input = editContainer?.querySelector(".report-name-input") as HTMLInputElement;
  
  if (displaySpan && editContainer && input) {
    displaySpan.style.display = "none";
    editContainer.style.display = "flex";
    input.focus();
    input.select();
  }
}

// 取消编辑报表名称
function cancelEditReportName(index: number): void {
  const tbody = document.getElementById("reports-table-body");
  if (!tbody) return;
  
  const displaySpan = tbody.querySelector(`.report-name-display[data-index="${index}"]`) as HTMLElement;
  const editContainer = displaySpan?.parentElement?.querySelector(".report-name-edit-container") as HTMLElement;
  
  if (displaySpan && editContainer) {
    displaySpan.style.display = "inline";
    editContainer.style.display = "none";
  }
}

// 确认编辑报表名称
async function confirmEditReportName(index: number, newName: string): Promise<void> {
  if (index < 0 || index >= reportsList.length) return;
  
  const oldName = reportsList[index].sheetName;
  newName = newName.trim();
  
  if (!newName || newName === oldName) {
    cancelEditReportName(index);
    return;
  }
  
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      const sheet = workbook.worksheets.getItem(oldName);
      
      // 修改 sheet 名称
      sheet.name = newName;
      
      // 同步修改报表标题（A1 单元格）
      const titleCell = sheet.getRange("A1");
      titleCell.values = [[newName]];
      
      await context.sync();
      
      // 更新本地数据
      reportsList[index].sheetName = newName;
      
      // 重新渲染表格
      updateReportsTable();
      
      showMessage(`报表名称已修改为: ${newName}`);
    });
  } catch (error) {
    console.error("修改报表名称失败:", error);
    showMessage(`修改名称失败: ${error.message}`, true);
    cancelEditReportName(index);
  }
}

// 跳转到指定工作表
async function navigateToSheet(sheetName: string): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      const sheet = workbook.worksheets.getItem(sheetName);
      sheet.activate();
      await context.sync();
      showMessage(`已跳转到报表: ${sheetName}`);
    });
  } catch (error) {
    showMessage(`跳转失败: 找不到工作表 "${sheetName}"`, true);
  }
}

// 显示确认对话框
function showConfirmDialog(message: string): Promise<boolean> {
  return new Promise((resolve) => {
    const overlay = document.getElementById("confirm-dialog");
    const messageEl = document.getElementById("confirm-dialog-message");
    const okButton = document.getElementById("confirm-dialog-ok");
    const cancelButton = document.getElementById("confirm-dialog-cancel");
    
    if (!overlay || !messageEl || !okButton || !cancelButton) {
      resolve(false);
      return;
    }
    
    messageEl.textContent = message;
    overlay.classList.add("show");
    
    const cleanup = () => {
      overlay.classList.remove("show");
      okButton.removeEventListener("click", onOk);
      cancelButton.removeEventListener("click", onCancel);
    };
    
    const onOk = () => {
      cleanup();
      resolve(true);
    };
    
    const onCancel = () => {
      cleanup();
      resolve(false);
    };
    
    okButton.addEventListener("click", onOk);
    cancelButton.addEventListener("click", onCancel);
  });
}

// 删除报表
async function deleteReport(index: number, sheetName: string): Promise<void> {
  const confirmed = await showConfirmDialog(`确定要删除报表 "${sheetName}" 吗？此操作不可恢复。`);
  
  if (!confirmed) {
    return;
  }
  
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      try {
        const sheet = workbook.worksheets.getItem(sheetName);
        sheet.delete();
        await context.sync();
        
        // 从列表中移除
        reportsList.splice(index, 1);
        updateReportsTable();
        showMessage(`已删除报表: ${sheetName}`);
      } catch (error) {
        // 如果工作表不存在，也从列表中移除
        reportsList.splice(index, 1);
        updateReportsTable();
        showMessage(`报表 "${sheetName}" 不存在，已从列表中移除`);
      }
    });
  } catch (error) {
    showMessage(`删除失败: ${error.message}`, true);
  }
}

// 添加报表到列表
function addReportToList(sheetName: string, rowCount: number, totalAmount: number): void {
  const report: ReportInfo = {
    sheetName: sheetName,
    createTime: new Date().toISOString(),
    rowCount: rowCount,
    totalAmount: totalAmount
  };
  
  reportsList.unshift(report); // 添加到列表开头
  updateReportsTable();
}


// 更新筛选条件列表显示
function updateFilterConditionsList(): void {
  const container = document.getElementById("filter-conditions-list");
  if (!container) return;
  
  if (filterConditionsList.length === 0) {
    container.innerHTML = '<div class="no-filters">暂无保存的筛选条件</div>';
    return;
  }
  
  container.innerHTML = filterConditionsList.map((filter, index) => {
    const createTime = new Date(filter.createTime).toLocaleString("zh-CN", { 
      month: "2-digit", 
      day: "2-digit", 
      hour: "2-digit", 
      minute: "2-digit" 
    });
    
    // 获取筛选条件摘要
    const filterSummary = filter.filterSettings.map(f => f.headerText).join(", ");
    
    return `
      <div class="filter-condition-item">
        <div class="filter-condition-info">
          <span class="filter-name-display" data-index="${index}" title="双击修改名称 | ${filterSummary}">${filter.name}</span>
          <div class="filter-name-edit-container" style="display: none;">
            <input type="text" class="filter-name-input" data-index="${index}" value="${filter.name}">
            <i class="ms-Icon ms-Icon--ReturnKey edit-hint-icon" title="按回车确认"></i>
          </div>
          <span class="filter-condition-time">${createTime}</span>
        </div>
        <div class="filter-condition-actions">
          <button class="icon-button-small apply-filter-btn" data-index="${index}" data-sheet-name="${filter.sheetName}" title="应用此筛选条件">
            <i class="ms-Icon ms-Icon--Play"></i>
          </button>
          <button class="icon-button-small danger delete-filter-btn" data-index="${index}" title="删除此筛选条件">
            <i class="ms-Icon ms-Icon--Delete"></i>
          </button>
        </div>
      </div>
    `;
  }).join("");
  
  // 绑定双击编辑筛选条件名称事件
  container.querySelectorAll(".filter-name-display").forEach(span => {
    span.addEventListener("dblclick", (e) => {
      const target = e.target as HTMLElement;
      const index = parseInt(target.getAttribute("data-index") || "0");
      startEditFilterName(index);
    });
  });
  
  // 绑定输入框事件
  container.querySelectorAll(".filter-name-input").forEach(input => {
    // 回车确认
    input.addEventListener("keydown", (e) => {
      const keyEvent = e as KeyboardEvent;
      if (keyEvent.key === "Enter") {
        const target = e.target as HTMLInputElement;
        const index = parseInt(target.getAttribute("data-index") || "0");
        confirmEditFilterName(index, target.value);
      } else if (keyEvent.key === "Escape") {
        const target = e.target as HTMLInputElement;
        const index = parseInt(target.getAttribute("data-index") || "0");
        cancelEditFilterName(index);
      }
    });
    
    // 失去焦点取消编辑
    input.addEventListener("blur", (e) => {
      const target = e.target as HTMLInputElement;
      const index = parseInt(target.getAttribute("data-index") || "0");
      setTimeout(() => {
        cancelEditFilterName(index);
      }, 200);
    });
  });
  
  // 绑定应用筛选条件事件
  container.querySelectorAll(".apply-filter-btn").forEach(button => {
    button.addEventListener("click", (e) => {
      const btn = e.currentTarget as HTMLElement;
      const index = parseInt(btn.getAttribute("data-index") || "0");
      const sheetName = btn.getAttribute("data-sheet-name");
      if (sheetName) {
        applyFilterCondition(index, sheetName);
      }
    });
  });
  
  // 绑定删除筛选条件事件
  container.querySelectorAll(".delete-filter-btn").forEach(button => {
    button.addEventListener("click", (e) => {
      const btn = e.currentTarget as HTMLElement;
      const index = parseInt(btn.getAttribute("data-index") || "0");
      deleteFilterCondition(index);
    });
  });
}

// 开始编辑筛选条件名称
function startEditFilterName(index: number): void {
  const container = document.getElementById("filter-conditions-list");
  if (!container) return;
  
  const displaySpan = container.querySelector(`.filter-name-display[data-index="${index}"]`) as HTMLElement;
  const editContainer = displaySpan?.parentElement?.querySelector(".filter-name-edit-container") as HTMLElement;
  const input = editContainer?.querySelector(".filter-name-input") as HTMLInputElement;
  
  if (displaySpan && editContainer && input) {
    displaySpan.style.display = "none";
    editContainer.style.display = "flex";
    input.focus();
    input.select();
  }
}

// 取消编辑筛选条件名称
function cancelEditFilterName(index: number): void {
  const container = document.getElementById("filter-conditions-list");
  if (!container) return;
  
  const displaySpan = container.querySelector(`.filter-name-display[data-index="${index}"]`) as HTMLElement;
  const editContainer = displaySpan?.parentElement?.querySelector(".filter-name-edit-container") as HTMLElement;
  
  if (displaySpan && editContainer) {
    displaySpan.style.display = "inline";
    editContainer.style.display = "none";
  }
}

// 确认编辑筛选条件名称
function confirmEditFilterName(index: number, newName: string): void {
  if (index < 0 || index >= filterConditionsList.length) return;
  
  newName = newName.trim();
  const oldName = filterConditionsList[index].name;
  
  if (!newName || newName === oldName) {
    cancelEditFilterName(index);
    return;
  }
  
  // 更新本地数据
  filterConditionsList[index].name = newName;
  
  // 重新渲染列表
  updateFilterConditionsList();
  
  showMessage(`筛选条件名称已修改为: ${newName}`);
}

// 应用筛选条件（从保存的筛选条件加载到面板并应用）
async function applyFilterCondition(index: number, sheetName: string): Promise<void> {
  if (index < 0 || index >= filterConditionsList.length) {
    showMessage("筛选条件不存在", true);
    return;
  }
  
  const filterCondition = filterConditionsList[index];
  
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      
      // 切换到主报表工作表
      const sheet = workbook.worksheets.getItem(filterCondition.config.sheetName);
      sheet.activate();
      sheet.load("name");
      
      // 获取使用范围的所有数据
      const usedRange = sheet.getUsedRange();
      usedRange.load("values, rowCount, columnCount, address, rowIndex");
      await context.sync();
      
      if (!usedRange.values || usedRange.values.length === 0) {
        showMessage("工作表没有数据", true);
        return;
      }
      
      // usedRange.rowIndex 是 0-based，第一行是 0
      const startRow = usedRange.rowIndex + 1; // 转换为 1-based Excel 行号
      const endRow = startRow + usedRange.rowCount - 1;
      
      // 更新当前主报表配置
      currentMainReportConfig = {
        dataRange: usedRange.address,
        headerRow: startRow,
        snColumn: filterCondition.config.snColumn,
        amtColumn: filterCondition.config.amtColumn,
        sheetName: sheet.name
      };
      
      // 重新构建筛选面板数据
      const headerRow = usedRange.values[0];
      filterPanelFields = [];
      
      for (let colIdx = 0; colIdx < usedRange.columnCount; colIdx++) {
        const headerText = headerRow[colIdx] ? String(headerRow[colIdx]) : getColumnName(colIdx);
        const allValues = new Set<string>();
        
        for (let rowIdx = 1; rowIdx < usedRange.values.length; rowIdx++) {
          const cellValue = usedRange.values[rowIdx][colIdx];
          const valueStr = cellValue !== null && cellValue !== undefined ? String(cellValue) : "";
          if (valueStr !== "") {
            allValues.add(valueStr);
          }
        }
        
        const sortedValues = Array.from(allValues).sort((a, b) => a.localeCompare(b, "zh-CN"));
        
        // 查找保存的筛选条件中是否有这一列
        const savedFilter = filterCondition.filterSettings.find(f => f.columnIndex === colIdx);
        
        // 如果保存了这一列的筛选条件，使用保存的值；否则默认全选
        let selectedValues: Set<string>;
        if (savedFilter && savedFilter.isFiltered) {
          selectedValues = new Set<string>(savedFilter.filterValues.filter(v => allValues.has(v)));
        } else {
          selectedValues = new Set<string>(sortedValues);
        }
        
        filterPanelFields.push({
          columnIndex: colIdx,
          headerText: headerText,
          allValues: sortedValues,
          selectedValues: selectedValues
        });
      }
      
      // 渲染筛选面板
      renderFilterPanel();
      
      // 启用筛选按钮
      const applyBtn = document.getElementById("btn-apply-filter") as HTMLButtonElement;
      const clearBtn = document.getElementById("btn-clear-filter") as HTMLButtonElement;
      const saveBtn = document.getElementById("btn-save-filter") as HTMLButtonElement;
      const generateBtn = document.getElementById("btn-generate") as HTMLButtonElement;
      const resetBtn = document.getElementById("btn-reset-filter") as HTMLButtonElement;
      if (applyBtn) applyBtn.disabled = false;
      if (clearBtn) clearBtn.disabled = false;
      if (saveBtn) saveBtn.disabled = false;
      if (generateBtn) generateBtn.disabled = false;
      if (resetBtn) resetBtn.disabled = false;
      
      // 首先显示所有数据行（从标题行的下一行开始）
      const dataStartRow = startRow + 1;
      for (let excelRow = dataStartRow; excelRow <= endRow; excelRow++) {
        const row = sheet.getRange(`${excelRow}:${excelRow}`);
        row.rowHidden = false;
      }
      await context.sync();
      
      // 构建筛选条件 - 只记录有筛选的字段
      const filterEntries: Array<{ colIdx: number; allowedValues: Set<string>; headerText: string }> = [];
      
      for (const field of filterPanelFields) {
        if (field.selectedValues.size < field.allValues.length && field.selectedValues.size > 0) {
          filterEntries.push({
            colIdx: field.columnIndex,
            allowedValues: field.selectedValues,
            headerText: field.headerText
          });
        }
      }
      
      // 如果没有筛选条件，直接返回
      if (filterEntries.length === 0) {
        updateCurrentFilterDisplay([]);
        showMessage("已加载筛选条件到面板，当前无筛选");
        return;
      }
      
      // 遍历每一行数据，应用筛选
      let hiddenCount = 0;
      let visibleCount = 0;
      
      for (let valueIdx = 1; valueIdx < usedRange.values.length; valueIdx++) {
        const rowData = usedRange.values[valueIdx];
        const excelRowNum = startRow + valueIdx; // 对应的 Excel 行号
        let shouldHide = false;
        
        for (const filter of filterEntries) {
          if (filter.colIdx < rowData.length) {
            const cellValue = rowData[filter.colIdx];
            const cellStr = cellValue !== null && cellValue !== undefined ? String(cellValue) : "";
            
            if (!filter.allowedValues.has(cellStr)) {
              shouldHide = true;
              break;
            }
          }
        }
        
        if (shouldHide) {
          const row = sheet.getRange(`${excelRowNum}:${excelRowNum}`);
          row.rowHidden = true;
          hiddenCount++;
        } else {
          visibleCount++;
        }
      }
      
      await context.sync();
      
      // 构建筛选条件设置用于显示
      const filterSettings: ColumnFilterSetting[] = filterEntries.map(f => ({
        columnIndex: f.colIdx,
        columnName: getColumnName(f.colIdx),
        headerText: f.headerText,
        filterType: "values",
        filterValues: Array.from(f.allowedValues),
        isFiltered: true
      }));
      
      updateCurrentFilterDisplay(filterSettings);
      
      // 显示结果
      const filterSummary = filterEntries.map(f => `${f.headerText}(${f.allowedValues.size}项)`).join(", ");
      showMessage(`已应用筛选条件: ${filterSummary} | 可见${visibleCount}行, 隐藏${hiddenCount}行`);
    });
  } catch (error) {
    console.error("应用筛选条件时出错:", error);
    appendDebugLog(`应用筛选条件失败: ${error.message}`);
    showMessage(`应用筛选条件失败: ${error.message}`, true);
  }
}

// 更新当前筛选条件显示
function updateCurrentFilterDisplay(filterSettings: ColumnFilterSetting[]): void {
  const container = document.getElementById("current-filter-display");
  if (!container) return;
  
  const activeFilters = filterSettings.filter(f => f.isFiltered);
  
  if (activeFilters.length === 0) {
    container.innerHTML = '<div class="no-filters">当前未应用筛选条件</div>';
    return;
  }
  
  // 生成每个筛选条件的 HTML
  const filterItemsHtml = activeFilters.map((filter, index) => {
    const values = filter.filterValues;
    const totalCount = values.length;
    const MAX_DISPLAY = 3; // 最多显示3个值
    
    if (totalCount <= MAX_DISPLAY) {
      // 值不多，全部显示
      return `
        <div class="current-filter-item">
          <span class="filter-field">${filter.headerText}</span>
          <span class="filter-separator">：</span>
          <span class="filter-value">${values.join("、")}</span>
        </div>
      `;
    } else {
      // 值太多，显示前几个 + "等xxx个结果"
      const displayValues = values.slice(0, MAX_DISPLAY);
      const remainingCount = totalCount;
      const filterId = `filter-expand-${index}`;
      
      return `
        <div class="current-filter-item">
          <span class="filter-field">${filter.headerText}</span>
          <span class="filter-separator">：</span>
          <span class="filter-value-collapsed" id="${filterId}-collapsed">
            ${displayValues.join("、")}
            <a href="javascript:void(0)" class="filter-expand-link" onclick="expandFilterValues('${filterId}', ${index})">等${remainingCount}个结果</a>
          </span>
          <span class="filter-value-expanded" id="${filterId}-expanded" style="display: none;">
            ${values.join("、")}
            <a href="javascript:void(0)" class="filter-collapse-link" onclick="collapseFilterValues('${filterId}')">收起</a>
          </span>
        </div>
      `;
    }
  }).join("");
  
  container.innerHTML = `
    <div class="current-filter-title">当前应用的筛选条件：</div>
    <div class="current-filter-items">
      ${filterItemsHtml}
    </div>
  `;
}

// 展开筛选值
function expandFilterValues(filterId: string, _filterIndex: number): void {
  const collapsedEl = document.getElementById(`${filterId}-collapsed`);
  const expandedEl = document.getElementById(`${filterId}-expanded`);
  
  if (collapsedEl && expandedEl) {
    collapsedEl.style.display = "none";
    expandedEl.style.display = "inline";
  }
}

// 收起筛选值
function collapseFilterValues(filterId: string): void {
  const collapsedEl = document.getElementById(`${filterId}-collapsed`);
  const expandedEl = document.getElementById(`${filterId}-expanded`);
  
  if (collapsedEl && expandedEl) {
    collapsedEl.style.display = "inline";
    expandedEl.style.display = "none";
  }
}

// 将函数暴露到全局作用域，以便 onclick 调用
(window as any).expandFilterValues = expandFilterValues;
(window as any).collapseFilterValues = collapseFilterValues;

// 检测并更新当前筛选条件显示（从筛选面板获取，只显示用户选择的字段）
function refreshCurrentFilterDisplay(): void {
  const container = document.getElementById("current-filter-display");
  if (!container) return;
  
  // 从筛选面板字段中获取有筛选的字段
  const columnFilterSettings: ColumnFilterSetting[] = [];
  
  for (const field of filterPanelFields) {
    // 只显示有筛选的字段（选中的值少于全部值，且至少选中一个）
    if (field.selectedValues.size < field.allValues.length && field.selectedValues.size > 0) {
      const selectedValuesArray = Array.from(field.selectedValues);
      
      // 传递完整的选中值数组，让 updateCurrentFilterDisplay 处理显示逻辑
      columnFilterSettings.push({
        columnIndex: field.columnIndex,
        columnName: getColumnName(field.columnIndex),
        headerText: field.headerText,
        filterType: "values",
        filterValues: selectedValuesArray,  // 传递完整的值数组，不截断
        isFiltered: true
      });
    }
  }
  
  // 更新显示
  if (columnFilterSettings.length > 0) {
    updateCurrentFilterDisplay(columnFilterSettings);
  } else if (filterPanelFields.length > 0) {
    // 筛选面板已加载，但没有筛选条件
    container.innerHTML = '<div class="no-filters">当前未应用筛选条件（所有字段均为全选）</div>';
  } else {
    // 筛选面板未加载
    container.innerHTML = '<div class="no-filters">请先点击"启用筛选"加载字段</div>';
  }
}

// 删除筛选条件
function deleteFilterCondition(index: number): void {
  if (index >= 0 && index < filterConditionsList.length) {
    filterConditionsList.splice(index, 1);
    updateFilterConditionsList();
    showMessage("已删除筛选条件");
  }
}

// 启用筛选
async function enableFilter(): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      const activeSheet = workbook.worksheets.getActiveWorksheet();
      activeSheet.load("name");
      
      // 获取使用范围
      const usedRange = activeSheet.getUsedRange();
      usedRange.load("values, rowCount, columnCount, address");
      await context.sync();
      
      if (!usedRange.values || usedRange.values.length < 2) {
        showMessage("当前工作表数据不足", true);
        return;
      }
      
      // 获取标题行（第一行）
      const headerRow = usedRange.values[0];
      
      // 更新当前主报表配置
      currentMainReportConfig = {
        dataRange: usedRange.address,
        headerRow: 1,
        snColumn: currentMainReportConfig?.snColumn || 0,
        amtColumn: currentMainReportConfig?.amtColumn || 0,
        sheetName: activeSheet.name
      };
      
      // 清空并重新构建筛选面板数据
      filterPanelFields = [];
      
      // 遍历每一列，收集唯一值
      for (let colIdx = 0; colIdx < usedRange.columnCount; colIdx++) {
        const headerText = headerRow[colIdx] ? String(headerRow[colIdx]) : getColumnName(colIdx);
        const allValues = new Set<string>();
        
        // 从第二行开始收集值
        for (let rowIdx = 1; rowIdx < usedRange.values.length; rowIdx++) {
          const cellValue = usedRange.values[rowIdx][colIdx];
          const valueStr = cellValue !== null && cellValue !== undefined ? String(cellValue) : "";
          if (valueStr !== "") {
            allValues.add(valueStr);
          }
        }
        
        // 将 Set 转换为排序后的数组
        const sortedValues = Array.from(allValues).sort((a, b) => a.localeCompare(b, "zh-CN"));
        
        filterPanelFields.push({
          columnIndex: colIdx,
          headerText: headerText,
          allValues: sortedValues,
          selectedValues: new Set<string>(sortedValues) // 默认全选
        });
      }
      
      // 渲染筛选面板
      renderFilterPanel();
      
      // 启用筛选按钮
      const applyBtn = document.getElementById("btn-apply-filter") as HTMLButtonElement;
      const clearBtn = document.getElementById("btn-clear-filter") as HTMLButtonElement;
      const saveBtn = document.getElementById("btn-save-filter") as HTMLButtonElement;
      const generateBtn = document.getElementById("btn-generate") as HTMLButtonElement;
      const resetBtn = document.getElementById("btn-reset-filter") as HTMLButtonElement;
      if (applyBtn) applyBtn.disabled = false;
      if (clearBtn) clearBtn.disabled = false;
      if (saveBtn) saveBtn.disabled = false;
      if (generateBtn) generateBtn.disabled = false;
      if (resetBtn) resetBtn.disabled = false;
      
      // 更新配置显示
      updateMainReportConfigDisplay();
      
      showMessage(`已加载 ${filterPanelFields.length} 个字段的筛选选项`);
    });
  } catch (error) {
    console.error("加载筛选面板时出错:", error);
    showMessage(`加载筛选面板失败: ${error.message}`, true);
  }
}

// 渲染筛选面板
// 当前激活的筛选字段索引
let activeFilterFieldIndex: number = -1;

function renderFilterPanel(): void {
  const panel = document.getElementById("filter-panel");
  if (!panel) return;
  
  if (filterPanelFields.length === 0) {
    panel.innerHTML = '<div class="filter-panel-hint">请先点击"启用筛选"加载字段</div>';
    return;
  }
  
  // 生成标签容器
  let tagsHtml = '<div class="filter-tags-container">';
  
  for (let i = 0; i < filterPanelFields.length; i++) {
    const field = filterPanelFields[i];
    const selectedCount = field.selectedValues.size;
    const totalCount = field.allValues.length;
    const isFiltered = selectedCount < totalCount;
    const isActive = i === activeFilterFieldIndex;
    
    let tagClass = "filter-tag";
    if (isFiltered) tagClass += " filtered";
    if (isActive) tagClass += " active";
    
    tagsHtml += `
      <div class="${tagClass}" data-field-index="${i}">
        <span class="filter-tag-name">${escapeHtml(field.headerText)}</span>
        <span class="filter-tag-count">${isFiltered ? selectedCount + '/' : ''}${totalCount}</span>
      </div>
    `;
  }
  
  tagsHtml += '</div>';
  
  // 生成下拉筛选区域（仅当有激活的字段时）
  let dropdownHtml = '';
  
  if (activeFilterFieldIndex >= 0 && activeFilterFieldIndex < filterPanelFields.length) {
    const field = filterPanelFields[activeFilterFieldIndex];
    const fieldId = `filter-field-${activeFilterFieldIndex}`;
    
    dropdownHtml = `
      <div class="filter-dropdown show" data-field-index="${activeFilterFieldIndex}">
        <div class="filter-dropdown-header">
          <span class="filter-dropdown-title">${escapeHtml(field.headerText)} 筛选</span>
          <button class="filter-dropdown-close" id="btn-close-dropdown">收起</button>
        </div>
        <input type="text" class="filter-field-search" placeholder="搜索选项..." 
               id="${fieldId}-search" data-field-index="${activeFilterFieldIndex}">
        <div class="filter-options-container" id="${fieldId}-options">
    `;
    
    for (let j = 0; j < field.allValues.length; j++) {
      const value = field.allValues[j];
      const isSelected = field.selectedValues.has(value);
      const displayValue = value.length > 40 ? value.substring(0, 40) + "..." : value;
      
      dropdownHtml += `
        <div class="filter-option ${isSelected ? 'selected' : ''}" 
             data-field-index="${activeFilterFieldIndex}" data-value-index="${j}" data-value="${escapeHtml(value)}">
          <input type="checkbox" ${isSelected ? 'checked' : ''} 
                 id="${fieldId}-opt-${j}" data-field-index="${activeFilterFieldIndex}" data-value="${escapeHtml(value)}">
          <span class="filter-option-label" title="${escapeHtml(value)}">${escapeHtml(displayValue)}</span>
        </div>
      `;
    }
    
    dropdownHtml += `
        </div>
        <div class="filter-select-actions" id="${fieldId}-actions">
          <button class="filter-select-action" data-field-index="${activeFilterFieldIndex}" data-action="all">全选</button>
          <button class="filter-select-action" data-field-index="${activeFilterFieldIndex}" data-action="none">全不选</button>
          <button class="filter-select-action" data-field-index="${activeFilterFieldIndex}" data-action="invert">反选</button>
        </div>
        <div class="filter-search-hint" id="${fieldId}-search-hint" style="display: none;">
          <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i>
          <span>以上操作仅针对当前搜索结果</span>
        </div>
      </div>
    `;
  }
  
  panel.innerHTML = tagsHtml + dropdownHtml;
  
  // 注意：不再每次都绑定事件，事件已在初始化时通过事件委托绑定到 panel 上
}

// HTML 转义函数
function escapeHtml(text: string): string {
  const div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}

// 标记：筛选面板事件是否已绑定
let filterPanelEventsBound = false;

// 绑定筛选面板事件（使用事件委托，只绑定一次）
function bindFilterPanelEvents(): void {
  const panel = document.getElementById("filter-panel");
  if (!panel || filterPanelEventsBound) return;
  
  filterPanelEventsBound = true;
  console.log("[筛选面板] 绑定事件委托");
  
  // 使用事件委托处理所有点击事件
  panel.addEventListener("click", (e) => {
    const target = e.target as HTMLElement;
    
    // 1. 标签点击事件 - 切换激活的字段
    const tagElement = target.closest(".filter-tag") as HTMLElement;
    if (tagElement) {
      const fieldIndex = parseInt(tagElement.getAttribute("data-field-index") || "-1", 10);
      
      if (fieldIndex === activeFilterFieldIndex) {
        activeFilterFieldIndex = -1;
      } else {
        activeFilterFieldIndex = fieldIndex;
      }
      
      renderFilterPanel();
      return;
    }
    
    // 2. 关闭下拉框按钮
    if (target.id === "btn-close-dropdown" || target.closest("#btn-close-dropdown")) {
      activeFilterFieldIndex = -1;
      renderFilterPanel();
      return;
    }
    
    // 3. 全选/全不选/反选按钮事件
    const actionButton = target.closest(".filter-select-action") as HTMLElement;
    if (actionButton) {
      e.preventDefault();
      e.stopPropagation();
      
      const fieldIndex = parseInt(actionButton.getAttribute("data-field-index") || "0", 10);
      const action = actionButton.getAttribute("data-action");
      
      console.log(`[筛选操作] action=${action}, fieldIndex=${fieldIndex}`);
      
      if (fieldIndex >= 0 && fieldIndex < filterPanelFields.length) {
        const field = filterPanelFields[fieldIndex];
        
        // 直接检查搜索框内容来判断是否处于搜索模式
        const searchInput = document.getElementById(`filter-field-${fieldIndex}-search`) as HTMLInputElement;
        const searchText = searchInput ? searchInput.value.trim() : "";
        const isSearchMode = searchText.length > 0;
        
        console.log(`[筛选操作] isSearchMode=${isSearchMode}, searchText="${searchText}"`);
        console.log(`[筛选操作] 操作前 selectedValues.size=${field.selectedValues.size}, allValues.length=${field.allValues.length}`);
        
        if (isSearchMode) {
          // 搜索模式：只操作可见选项
          const optionsContainer = document.getElementById(`filter-field-${fieldIndex}-options`);
          const visibleValues: string[] = [];
          if (optionsContainer) {
            const options = optionsContainer.querySelectorAll(".filter-option:not(.hidden)");
            options.forEach((option) => {
              const valueIndexStr = option.getAttribute("data-value-index");
              const valueIndex = valueIndexStr !== null ? parseInt(valueIndexStr, 10) : -1;
              const value = valueIndex >= 0 && valueIndex < field.allValues.length 
                ? field.allValues[valueIndex] 
                : option.getAttribute("data-value") || "";
              visibleValues.push(value);
            });
          }
          
          console.log(`[筛选操作] 搜索模式，可见选项数: ${visibleValues.length}`);
          
          if (action === "all") {
            for (const v of visibleValues) {
              field.selectedValues.add(v);
            }
          } else if (action === "none") {
            for (const v of visibleValues) {
              field.selectedValues.delete(v);
            }
          } else if (action === "invert") {
            for (const v of visibleValues) {
              if (field.selectedValues.has(v)) {
                field.selectedValues.delete(v);
              } else {
                field.selectedValues.add(v);
              }
            }
          }
        } else {
          // 非搜索模式：操作所有选项
          console.log(`[筛选操作] 非搜索模式，操作所有 ${field.allValues.length} 个选项`);
          
          if (action === "all") {
            field.selectedValues = new Set<string>(field.allValues);
          } else if (action === "none") {
            field.selectedValues = new Set<string>();
          } else if (action === "invert") {
            const newSelected = new Set<string>();
            for (const v of field.allValues) {
              if (!field.selectedValues.has(v)) {
                newSelected.add(v);
              }
            }
            console.log(`[筛选操作] 反选：之前选中 ${field.selectedValues.size} 个，反选后应选中 ${newSelected.size} 个`);
            field.selectedValues = newSelected;
          }
        }
        
        console.log(`[筛选操作] 操作后 selectedValues.size=${field.selectedValues.size}`);
        
        // 重新渲染下拉框中的选项
        updateFieldOptions(fieldIndex);
        // 更新标签上的计数
        updateTagCount(fieldIndex);
      }
      return;
    }
    
    // 4. 点击选项行切换选中状态（但不是复选框本身）
    const optionDiv = target.closest(".filter-option") as HTMLElement;
    if (optionDiv && target.tagName !== "INPUT") {
      const fieldIndex = parseInt(optionDiv.getAttribute("data-field-index") || "0", 10);
      const checkbox = optionDiv.querySelector("input[type='checkbox']") as HTMLInputElement;
      
      if (fieldIndex >= 0 && fieldIndex < filterPanelFields.length && checkbox) {
        const field = filterPanelFields[fieldIndex];
        
        // 使用 data-value-index 来获取原始值
        const valueIndexStr = optionDiv.getAttribute("data-value-index");
        const valueIndex = valueIndexStr !== null ? parseInt(valueIndexStr, 10) : -1;
        const value = valueIndex >= 0 && valueIndex < field.allValues.length 
          ? field.allValues[valueIndex] 
          : optionDiv.getAttribute("data-value") || "";
        
        // 切换选中状态
        const newChecked = !checkbox.checked;
        checkbox.checked = newChecked;
        
        if (newChecked) {
          field.selectedValues.add(value);
        } else {
          field.selectedValues.delete(value);
        }
        
        // 更新选项的 selected 样式
        optionDiv.classList.toggle("selected", newChecked);
        
        // 更新标签上的计数
        updateTagCount(fieldIndex);
      }
      return;
    }
  });
  
  // 搜索框输入事件
  panel.addEventListener("input", (e) => {
    const target = e.target as HTMLInputElement;
    if (target.classList.contains("filter-field-search")) {
      const fieldIndex = parseInt(target.getAttribute("data-field-index") || "0", 10);
      filterSearchOptions(fieldIndex, target.value);
    }
  });
  
  // 复选框 change 事件（用于直接点击复选框的情况）
  panel.addEventListener("change", (e) => {
    const target = e.target as HTMLInputElement;
    if (target.type === "checkbox" && target.closest(".filter-option")) {
      const fieldIndex = parseInt(target.getAttribute("data-field-index") || "0", 10);
      
      if (fieldIndex >= 0 && fieldIndex < filterPanelFields.length) {
        const field = filterPanelFields[fieldIndex];
        
        // 使用 data-value-index 来获取原始值
        const optionDiv = target.closest(".filter-option") as HTMLElement;
        let value = target.getAttribute("data-value") || "";
        if (optionDiv) {
          const valueIndexStr = optionDiv.getAttribute("data-value-index");
          const valueIndex = valueIndexStr !== null ? parseInt(valueIndexStr, 10) : -1;
          if (valueIndex >= 0 && valueIndex < field.allValues.length) {
            value = field.allValues[valueIndex];
          }
        }
        
        if (target.checked) {
          field.selectedValues.add(value);
        } else {
          field.selectedValues.delete(value);
        }
        
        // 更新选项的 selected 样式
        if (optionDiv) {
          optionDiv.classList.toggle("selected", target.checked);
        }
        
        // 更新标签上的计数
        updateTagCount(fieldIndex);
      }
    }
  });
}

// 更新标签上的计数显示
function updateTagCount(fieldIndex: number): void {
  const field = filterPanelFields[fieldIndex];
  if (!field) return;
  
  const tag = document.querySelector(`.filter-tag[data-field-index="${fieldIndex}"]`);
  if (!tag) return;
  
  const selectedCount = field.selectedValues.size;
  const totalCount = field.allValues.length;
  const isFiltered = selectedCount < totalCount;
  
  // 更新标签样式
  tag.classList.toggle("filtered", isFiltered);
  
  // 更新计数显示
  const countSpan = tag.querySelector(".filter-tag-count");
  if (countSpan) {
    countSpan.textContent = isFiltered ? `${selectedCount}/${totalCount}` : `${totalCount}`;
  }
}

// 根据搜索关键字过滤选项
function filterSearchOptions(fieldIndex: number, searchText: string): void {
  const optionsContainer = document.getElementById(`filter-field-${fieldIndex}-options`);
  const actionsContainer = document.getElementById(`filter-field-${fieldIndex}-actions`);
  const searchHint = document.getElementById(`filter-field-${fieldIndex}-search-hint`);
  if (!optionsContainer) return;
  
  const options = optionsContainer.querySelectorAll(".filter-option");
  const searchLower = searchText.toLowerCase().trim();
  const hasSearchText = searchLower.length > 0;
  
  let visibleCount = 0;
  options.forEach((option) => {
    const value = option.getAttribute("data-value") || "";
    const matches = value.toLowerCase().includes(searchLower);
    option.classList.toggle("hidden", !matches);
    if (matches) visibleCount++;
  });
  
  // 当有搜索文字时，显示提示并给按钮添加特殊样式
  if (actionsContainer) {
    actionsContainer.classList.toggle("search-mode", hasSearchText);
  }
  if (searchHint) {
    searchHint.style.display = hasSearchText ? "flex" : "none";
    // 更新提示文字，显示当前可见的选项数量
    const hintSpan = searchHint.querySelector("span");
    if (hintSpan && hasSearchText) {
      hintSpan.textContent = `以上操作仅针对当前搜索结果（${visibleCount} 项）`;
    }
  }
}

// 更新字段的选项（用于全选/反选后）
function updateFieldOptions(fieldIndex: number): void {
  const field = filterPanelFields[fieldIndex];
  const optionsContainer = document.getElementById(`filter-field-${fieldIndex}-options`);
  if (!optionsContainer || !field) return;
  
  const options = optionsContainer.querySelectorAll(".filter-option");
  options.forEach((option) => {
    // 使用 data-value-index 来获取原始值，避免 HTML 编码问题
    const valueIndexStr = option.getAttribute("data-value-index");
    const valueIndex = valueIndexStr !== null ? parseInt(valueIndexStr, 10) : -1;
    const value = valueIndex >= 0 && valueIndex < field.allValues.length 
      ? field.allValues[valueIndex] 
      : option.getAttribute("data-value") || "";
    
    const isSelected = field.selectedValues.has(value);
    const checkbox = option.querySelector("input[type='checkbox']") as HTMLInputElement;
    if (checkbox) {
      checkbox.checked = isSelected;
    }
    option.classList.toggle("selected", isSelected);
  });
  
  console.log(`[updateFieldOptions] fieldIndex=${fieldIndex}, 更新了 ${options.length} 个选项`);
}

// 更新字段计数显示
function updateFieldCountDisplay(fieldIndex: number): void {
  const field = filterPanelFields[fieldIndex];
  const fieldGroup = document.querySelector(`[data-field-index="${fieldIndex}"].filter-field-group`);
  if (!fieldGroup || !field) return;
  
  const label = fieldGroup.querySelector(".filter-field-label");
  if (label) {
    const selectedCount = field.selectedValues.size;
    const totalCount = field.allValues.length;
    const isFiltered = selectedCount < totalCount;
    
    label.innerHTML = `
      ${field.headerText} 
      <span style="font-weight: normal; color: ${isFiltered ? '#0078d4' : '#605e5c'};">
        (${selectedCount}/${totalCount})
      </span>
    `;
  }
}

// 应用筛选条件到 Excel
async function applyFilterFromPanel(): Promise<void> {
  if (filterPanelFields.length === 0) {
    showMessage("请先加载筛选面板", true);
    return;
  }
  
  // 立即显示处理提示
  showMessage("正在应用筛选条件...", false);
  
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      let targetSheet: Excel.Worksheet;
      
      if (currentMainReportConfig) {
        targetSheet = workbook.worksheets.getItem(currentMainReportConfig.sheetName);
      } else {
        targetSheet = workbook.worksheets.getActiveWorksheet();
      }
      
      targetSheet.activate();
      
      // 获取使用范围
      const usedRange = targetSheet.getUsedRange();
      usedRange.load("values, rowCount, rowIndex, columnCount");
      await context.sync();
      
      if (!usedRange.values || usedRange.values.length === 0) {
        showMessage("工作表没有数据", true);
        return;
      }
      
      const startRow = usedRange.rowIndex; // 0-based
      const rowCount = usedRange.rowCount;
      const columnCount = usedRange.columnCount;
      
      // 构建筛选条件
      const filterEntries: Array<{ colIdx: number; allowedValues: Set<string>; headerText: string }> = [];
      
      for (const field of filterPanelFields) {
        if (field.selectedValues.size < field.allValues.length && field.selectedValues.size > 0) {
          filterEntries.push({
            colIdx: field.columnIndex,
            allowedValues: field.selectedValues,
            headerText: field.headerText
          });
        }
      }
      
      // ===== 优化算法：在内存中计算每行的可见状态 =====
      const rowVisibility: boolean[] = []; // true = 可见, false = 隐藏
      rowVisibility.push(true); // 标题行始终可见
      
      let visibleCount = 0;
      let hiddenCount = 0;
      let totalAmount = 0; // 合计金额
      
      // 获取金额列索引（如果配置了的话）
      const amtColIndex = currentMainReportConfig?.amtColumn ?? -1;
      
      // 在内存中判断每行是否应该显示
      for (let valueIdx = 1; valueIdx < usedRange.values.length; valueIdx++) {
        const rowData = usedRange.values[valueIdx];
        let shouldShow = true;
        
        if (filterEntries.length > 0) {
          for (const filter of filterEntries) {
            if (filter.colIdx < rowData.length) {
              const cellValue = rowData[filter.colIdx];
              const cellStr = cellValue !== null && cellValue !== undefined ? String(cellValue) : "";
              
              if (!filter.allowedValues.has(cellStr)) {
                shouldShow = false;
                break;
              }
            }
          }
        }
        
        rowVisibility.push(shouldShow);
        if (shouldShow) {
          visibleCount++;
          // 累加可见行的金额
          if (amtColIndex >= 0 && amtColIndex < rowData.length) {
            const amountValue = rowData[amtColIndex];
            totalAmount += cleanAmount(amountValue);
          }
        } else {
          hiddenCount++;
        }
      }
      
      // 更新筛选统计数据
      currentFilterStatistics = {
        filteredRowCount: visibleCount,
        totalAmount: totalAmount,
        isValid: true
      };
      
      // ===== 批量设置行可见性 =====
      // 先重置所有数据行为可见
      const allDataRowsStart = startRow + 2; // Excel 行号，跳过标题行
      const allDataRowsEnd = startRow + rowCount;
      if (allDataRowsEnd >= allDataRowsStart) {
        const allDataRows = targetSheet.getRange(`${allDataRowsStart}:${allDataRowsEnd}`);
        allDataRows.rowHidden = false;
      }
      await context.sync();
      
      // 然后设置需要隐藏的行
      // 将连续的隐藏行合并为一个范围操作
      let i = 1; // 从数据行开始（跳过标题行）
      let hiddenRangeCount = 0;
      while (i < rowVisibility.length) {
        if (!rowVisibility[i]) {
          // 找到需要隐藏的起始行
          let endIdx = i;
          
          // 查找连续需要隐藏的行
          while (endIdx + 1 < rowVisibility.length && !rowVisibility[endIdx + 1]) {
            endIdx++;
          }
          
          // 批量隐藏这个范围
          const rangeStartRow = startRow + i + 1; // Excel 行号 (1-based)
          const rangeEndRow = startRow + endIdx + 1;
          const range = targetSheet.getRange(`${rangeStartRow}:${rangeEndRow}`);
          range.rowHidden = true;
          hiddenRangeCount++;
          
          i = endIdx + 1;
        } else {
          i++;
        }
      }
      
      await context.sync();
      
      // 调试信息
      appendDebugLog(`应用筛选完成: 隐藏了 ${hiddenRangeCount} 个连续范围, 共 ${hiddenCount} 行`);
      appendDebugLog(`数据起始行: Excel 第 ${startRow + 1} 行, 共 ${rowCount} 行`);
      
      // ===== 统一可见行的行高 =====
      // 获取所有数据行，设置统一行高（仅对可见行生效）
      const dataRange = targetSheet.getRangeByIndexes(startRow + 1, 0, rowCount - 1, columnCount);
      dataRange.format.rowHeight = 20; // 设置统一行高为 20 像素
      await context.sync();
      
      // 构建筛选条件设置用于显示
      const filterSettings: ColumnFilterSetting[] = filterEntries.map(f => ({
        columnIndex: f.colIdx,
        columnName: getColumnName(f.colIdx),
        headerText: f.headerText,
        filterType: "values",
        filterValues: Array.from(f.allowedValues),
        isFiltered: true
      }));
      
      updateCurrentFilterDisplay(filterSettings);
      
      // 更新筛选条件文本区域并自动复制到剪贴板
      updateFilterTextDisplay();
      await copyFilterTextToClipboard();
      
      // 显示结果
      const filterSummary = filterEntries.map(f => {
        const count = f.allowedValues.size;
        return `${f.headerText}(${count}项)`;
      }).join(", ");
      
      showMessage(`已应用筛选: ${filterSummary} | 可见${visibleCount}行, 隐藏${hiddenCount}行`);
      appendDebugLog(`应用筛选: ${filterSummary}`);
    });
  } catch (error) {
    console.error("应用筛选时出错:", error);
    showMessage(`应用筛选失败: ${error.message}`, true);
  }
}

// 清除筛选条件
async function clearFilterFromPanel(): Promise<void> {
  // 重置所有字段为全选
  for (const field of filterPanelFields) {
    field.selectedValues = new Set<string>(field.allValues);
  }
  
  // 重置筛选统计数据
  currentFilterStatistics = {
    filteredRowCount: 0,
    totalAmount: 0,
    isValid: false
  };
  
  // 重新渲染面板
  renderFilterPanel();
  
  // 应用（显示所有行）
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      let targetSheet: Excel.Worksheet;
      
      if (currentMainReportConfig) {
        targetSheet = workbook.worksheets.getItem(currentMainReportConfig.sheetName);
      } else {
        targetSheet = workbook.worksheets.getActiveWorksheet();
      }
      
      const usedRange = targetSheet.getUsedRange();
      usedRange.load("rowCount, rowIndex, columnCount");
      await context.sync();
      
      // 优化：一次性显示所有数据行（跳过标题行）
      const dataRowCount = usedRange.rowCount - 1;
      if (dataRowCount > 0) {
        const startRowNum = usedRange.rowIndex + 2; // Excel 行号，跳过标题行
        const endRowNum = usedRange.rowIndex + usedRange.rowCount;
        const allDataRows = targetSheet.getRange(`${startRowNum}:${endRowNum}`);
        allDataRows.rowHidden = false;
        
        // 统一行高
        allDataRows.format.rowHeight = 20;
      }
      await context.sync();
      
      updateCurrentFilterDisplay([]);
      
      // 隐藏筛选条件文本区域
      const filterTextContainer = document.getElementById("filter-text-container");
      if (filterTextContainer) {
        filterTextContainer.style.display = "none";
      }
      
      showMessage("已清除所有筛选");
    });
  } catch (error) {
    console.error("清除筛选时出错:", error);
    showMessage(`清除筛选失败: ${error.message}`, true);
  }
}

// 保存当前筛选条件（从面板）
function saveFilterFromPanel(): void {
  if (filterPanelFields.length === 0) {
    showMessage("请先加载筛选面板", true);
    return;
  }
  
  // 构建筛选条件设置
  const filterSettings: ColumnFilterSetting[] = [];
  
  for (const field of filterPanelFields) {
    if (field.selectedValues.size < field.allValues.length && field.selectedValues.size > 0) {
      filterSettings.push({
        columnIndex: field.columnIndex,
        columnName: getColumnName(field.columnIndex),
        headerText: field.headerText,
        filterType: "values",
        filterValues: Array.from(field.selectedValues),
        isFiltered: true
      });
    }
  }
  
  if (filterSettings.length === 0) {
    showMessage("没有需要保存的筛选条件（所有字段都是全选状态）", true);
    return;
  }
  
  // 创建筛选条件记录
  const filterCondition: FilterCondition = {
    id: `filter_${Date.now()}`,
    name: `筛选_${new Date().toLocaleString("zh-CN", { month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit" })}`,
    createTime: new Date().toISOString(),
    sheetName: currentMainReportConfig?.sheetName || "未知",
    filterSettings: filterSettings,
    config: currentMainReportConfig ? { ...currentMainReportConfig } : {
      dataRange: "",
      headerRow: 1,
      snColumn: 0,
      amtColumn: 0,
      sheetName: ""
    }
  };
  
  filterConditionsList.unshift(filterCondition);
  updateFilterConditionsList();
  
  const filterSummary = filterSettings.map(f => `${f.headerText}(${f.filterValues.length}项)`).join(", ");
  showMessage(`已保存筛选条件: ${filterSummary}`);
}

// 保存当前筛选条件（记录每列的筛选条件详情）
async function saveCurrentFilterCondition(): Promise<void> {
  // 从筛选面板获取筛选条件（只记录用户选择的字段）
  const columnFilterSettings: ColumnFilterSetting[] = [];
  
  for (const field of filterPanelFields) {
    // 只记录有筛选的字段（选中的值少于全部值，且至少选中一个）
    if (field.selectedValues.size < field.allValues.length && field.selectedValues.size > 0) {
      columnFilterSettings.push({
        columnIndex: field.columnIndex,
        columnName: getColumnName(field.columnIndex),
        headerText: field.headerText,
        filterType: "values",
        filterValues: Array.from(field.selectedValues),
        isFiltered: true
      });
      
      appendDebugLog(`筛选条件 - ${field.headerText}: ${field.selectedValues.size}/${field.allValues.length}项`);
    }
  }
  
  // 如果没有筛选条件，不保存
  if (columnFilterSettings.length === 0) {
    appendDebugLog("没有筛选条件需要保存");
    return;
  }
  
  // 创建筛选条件记录
  const filterCondition: FilterCondition = {
    id: `filter_${Date.now()}`,
    name: `筛选条件_${new Date().toLocaleString("zh-CN", { month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit" })}`,
    createTime: new Date().toISOString(),
    sheetName: currentMainReportConfig?.sheetName || "未知",
    filterSettings: columnFilterSettings,
    config: currentMainReportConfig ? { ...currentMainReportConfig } : {
      dataRange: "",
      headerRow: 1,
      snColumn: 0,
      amtColumn: 0,
      sheetName: ""
    }
  };
  
  filterConditionsList.unshift(filterCondition);
  updateFilterConditionsList();
  refreshCurrentFilterDisplay();
  
  const filterCount = columnFilterSettings.length;
  appendDebugLog(`已保存筛选条件: ${filterCount} 列有筛选`);
}

// 跳转到主报表的第一行第一个单元格
async function goToMainReport(): Promise<void> {
  if (!currentMainReportConfig || !currentMainReportConfig.sheetName) {
    showMessage("未配置主报表", true);
    return;
  }
  
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      const sheet = workbook.worksheets.getItem(currentMainReportConfig!.sheetName);
      sheet.activate();
      
      // 选中第一行第一个单元格
      const firstCell = sheet.getRange("A1");
      firstCell.select();
      
      await context.sync();
      showMessage(`已跳转到主报表: ${currentMainReportConfig!.sheetName}`);
    });
  } catch (error) {
    console.error("跳转到主报表失败:", error);
    showMessage(`跳转失败: ${error.message}`, true);
  }
}

// 重置筛选并跳转到主报表
async function resetFilterAndGoToMainReport(): Promise<void> {
  if (!currentMainReportConfig || !currentMainReportConfig.sheetName) {
    showMessage("未配置主报表", true);
    return;
  }
  
  // 重置筛选统计数据
  currentFilterStatistics = {
    filteredRowCount: 0,
    totalAmount: 0,
    isValid: false
  };
  
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      const sheet = workbook.worksheets.getItem(currentMainReportConfig!.sheetName);
      sheet.activate();
      
      // 获取使用范围
      const usedRange = sheet.getUsedRange();
      usedRange.load("rowCount, rowIndex");
      await context.sync();
      
      // 优化：一次性显示所有数据行
      const dataRowCount = usedRange.rowCount - 1;
      if (dataRowCount > 0) {
        const startRowNum = usedRange.rowIndex + 2; // 跳过标题行
        const endRowNum = usedRange.rowIndex + usedRange.rowCount;
        const allDataRows = sheet.getRange(`${startRowNum}:${endRowNum}`);
        allDataRows.rowHidden = false;
        
        // 统一行高
        allDataRows.format.rowHeight = 20;
      }
      
      // 选中第一行第一个单元格
      const firstCell = sheet.getRange("A1");
      firstCell.select();
      
      await context.sync();
      
      // 重置筛选面板（全部字段设为全选）
      for (const field of filterPanelFields) {
        field.selectedValues = new Set<string>(field.allValues);
      }
      
      // 重新渲染筛选面板
      activeFilterFieldIndex = -1;
      renderFilterPanel();
      refreshCurrentFilterDisplay();
      
      // 隐藏筛选条件文本区域
      const filterTextContainer = document.getElementById("filter-text-container");
      if (filterTextContainer) {
        filterTextContainer.style.display = "none";
      }
      
      showMessage("已重置筛选并跳转到主报表");
    });
  } catch (error) {
    console.error("重置筛选失败:", error);
    showMessage(`重置失败: ${error.message}`, true);
  }
}

// 生成报表
async function generateReport(): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      
      // 使用主报表工作表，而不是当前活动工作表
      let sourceSheet: Excel.Worksheet;
      let usingMainReport = false;
      if (currentMainReportConfig && currentMainReportConfig.sheetName) {
        sourceSheet = workbook.worksheets.getItem(currentMainReportConfig.sheetName);
        usingMainReport = true;
      } else {
        sourceSheet = workbook.worksheets.getActiveWorksheet();
      }
      sourceSheet.load("name");
      await context.sync();
      
      // 调试：输出工作表信息
      appendDebugLog(`生成报表 - 使用工作表: "${sourceSheet.name}", 是否主报表配置: ${usingMainReport}`);
      if (currentMainReportConfig) {
        appendDebugLog(`  主报表配置: sheetName="${currentMainReportConfig.sheetName}"`);
      }
      
      // 检查配置是否存在
      let snColIndex: number | null = null;
      let amtColIndex: number | null = null;
      let sourceSheetName: string = ""; // 保存源表名称
      
      
      try {
        const snNamedItem = workbook.names.getItem(CFG_SN_COL_NAME);
        snNamedItem.load("name, formula");
        await context.sync();
        
        // 在第一次 sync 后保存源表名称
        sourceSheetName = sourceSheet.name;
        
        // 方法1：从 NamedItem 的 Range 获取列索引
        const snRange = snNamedItem.getRange();
        snRange.load("columnIndex, address, columnCount");
        await context.sync();
        let snColIndexFromRange = snRange.columnIndex;
        
        // 方法2：从公式中解析列名，然后转换为列索引（备用方法）
        const snColumnName = parseColumnFromFormula(snNamedItem.formula);
        let snColIndexFromFormula: number | null = null;
        if (snColumnName) {
          snColIndexFromFormula = getColumnIndexFromName(snColumnName);
        }
        
        // 使用两种方法中更可靠的一个
        snColIndex = snColIndexFromFormula !== null ? snColIndexFromFormula : snColIndexFromRange;
        
        // 调试信息：验证 S/N 列索引
        appendDebugLog(`S/N 列 NamedItem: ${snNamedItem.name}`);
        appendDebugLog(`  公式: ${snNamedItem.formula}`);
        appendDebugLog(`  地址: ${snRange.address}`);
        appendDebugLog(`  从 Range 获取列索引: ${snColIndexFromRange} (列 ${getColumnName(snColIndexFromRange)})`);
        if (snColumnName) {
          appendDebugLog(`  从公式解析列名: ${snColumnName}, 列索引: ${snColIndexFromFormula}`);
        }
        appendDebugLog(`  最终使用列索引: ${snColIndex} (列 ${getColumnName(snColIndex)})`);
      } catch (error) {
        showMessage("错误: 未配置 S/N 列，请先点击\"设为 S/N 列\"按钮", true);
        return;
      }
      
      try {
        const amtNamedItem = workbook.names.getItem(CFG_AMT_COL_NAME);
        amtNamedItem.load("name, formula");
        await context.sync();
        
        // 方法1：从 NamedItem 的 Range 获取列索引
        const amtRange = amtNamedItem.getRange();
        amtRange.load("columnIndex, address, columnCount");
        await context.sync();
        let amtColIndexFromRange = amtRange.columnIndex;
        
        // 方法2：从公式中解析列名，然后转换为列索引（备用方法）
        const amtColumnName = parseColumnFromFormula(amtNamedItem.formula);
        let amtColIndexFromFormula: number | null = null;
        if (amtColumnName) {
          amtColIndexFromFormula = getColumnIndexFromName(amtColumnName);
        }
        
        // 使用两种方法中更可靠的一个
        amtColIndex = amtColIndexFromFormula !== null ? amtColIndexFromFormula : amtColIndexFromRange;
        
        // 调试信息：验证金额列索引
        appendDebugLog(`金额列 NamedItem: ${amtNamedItem.name}`);
        appendDebugLog(`  公式: ${amtNamedItem.formula}`);
        appendDebugLog(`  地址: ${amtRange.address}`);
        appendDebugLog(`  从 Range 获取列索引: ${amtColIndexFromRange} (列 ${getColumnName(amtColIndexFromRange)})`);
        if (amtColumnName) {
          appendDebugLog(`  从公式解析列名: ${amtColumnName}, 列索引: ${amtColIndexFromFormula}`);
        }
        appendDebugLog(`  最终使用列索引: ${amtColIndex} (列 ${getColumnName(amtColIndex)})`);
      } catch (error) {
        showMessage("错误: 未配置金额列，请先点击\"设为金额列\"按钮", true);
        return;
      }
      
      // 读取筛选后的可见数据（用户已经在 Excel 中筛选好的数据）
      const usedRange = sourceSheet.getUsedRange();
      if (!usedRange) {
        showMessage("错误: 无法获取工作表使用范围", true);
        return;
      }
      
      usedRange.load("rowCount, columnCount, address, values, rowIndex");
      
      // 先获取使用范围信息
      await context.sync();
      
      if (usedRange.rowCount === 0 || usedRange.columnCount === 0) {
        showMessage("错误: 当前工作表没有数据", true);
        return;
      }
      
      const rowCount = usedRange.rowCount;
      const columnCount = usedRange.columnCount;
      
      // 验证数据已加载
      if (!usedRange.values || usedRange.values.length === 0) {
        showMessage(`错误: 无法读取工作表数据 (行数: ${rowCount}, 列数: ${columnCount})`, true);
        return;
      }
      
      // 清空调试日志
      clearDebugLog();
      appendDebugLog(`开始处理数据...`);
      appendDebugLog(`S/N 列索引: ${snColIndex}, 金额列索引: ${amtColIndex}`);
      
      // 验证列索引是否正确
      if (snColIndex === amtColIndex) {
        appendDebugLog(`警告: S/N 列和金额列的索引相同！这可能导致错误。`);
      }
      
      // ===== 直接使用 filterPanelFields 过滤数据（更快、更可靠）=====
      // 构建筛选条件
      const filterEntries: Array<{ colIdx: number; allowedValues: Set<string>; headerText: string }> = [];
      for (const field of filterPanelFields) {
        if (field.selectedValues.size < field.allValues.length && field.selectedValues.size > 0) {
          filterEntries.push({
            colIdx: field.columnIndex,
            allowedValues: field.selectedValues,
            headerText: field.headerText
          });
        }
      }
      
      appendDebugLog(`筛选条件数量: ${filterEntries.length}`);
      for (const filter of filterEntries) {
        appendDebugLog(`  ${filter.headerText}: ${filter.allowedValues.size} 项`);
      }
      
      // 在内存中处理数据
      const filteredRows: any[][] = [];
      
      // 确保每行数据的列数与 columnCount 一致的辅助函数
      const normalizeRow = (rowData: any[]): any[] => {
        if (!rowData || !Array.isArray(rowData)) {
          return new Array(columnCount).fill("");
        }
        if (rowData.length === columnCount) {
          return rowData;
        } else if (rowData.length < columnCount) {
          return [...rowData, ...new Array(columnCount - rowData.length).fill("")];
        } else {
          return rowData.slice(0, columnCount);
        }
      };
      
      // 第一行是表头，始终保留
      if (usedRange.values.length > 0 && usedRange.values[0]) {
        filteredRows.push(normalizeRow(usedRange.values[0]));
      }
      
      // 从第二行开始，应用筛选条件
      let visibleRowCount = 0;
      let hiddenRowCount = 0;
      
      for (let i = 1; i < usedRange.values.length; i++) {
        const rowData = usedRange.values[i];
        if (!rowData || !Array.isArray(rowData)) continue;
        
        // 检查该行是否符合所有筛选条件
        let shouldInclude = true;
        
        if (filterEntries.length > 0) {
          for (const filter of filterEntries) {
            if (filter.colIdx < rowData.length) {
              const cellValue = rowData[filter.colIdx];
              const cellStr = cellValue !== null && cellValue !== undefined ? String(cellValue) : "";
              
              if (!filter.allowedValues.has(cellStr)) {
                shouldInclude = false;
                break;
              }
            }
          }
        }
        
        if (shouldInclude) {
          filteredRows.push(normalizeRow(rowData));
          visibleRowCount++;
        } else {
          hiddenRowCount++;
        }
      }
      
      appendDebugLog(`总行数: ${rowCount}, 符合条件: ${visibleRowCount}, 不符合条件: ${hiddenRowCount}`);
      
      const filteredCount = filteredRows.length - 1; // 减去表头
      
      // 如果没有可见的数据行（只有表头），提示用户
      if (filteredCount === 0) {
        showMessage("筛选后没有符合条件的数据行，请调整筛选条件后重试。", true);
        return;
      }
      appendDebugLog(`筛选完成，可见数据行数: ${filteredCount}`);
      
      // 从 filteredRows 计算总金额（只累加数据行，跳过表头）
      let totalAmount = 0;
      const amountValues: number[] = []; // 用于调试
      
      for (let i = 1; i < filteredRows.length; i++) {
        const rowData = filteredRows[i];
        // Excel 行号 = filteredRows 索引 + 1（因为 filteredRows[0] 是表头）
        const excelRowNumber = i + 1;
        
        if (rowData && Array.isArray(rowData) && amtColIndex !== null && amtColIndex >= 0 && amtColIndex < rowData.length) {
          const amountValue = rowData[amtColIndex];
          
          // 验证：同时显示 S/N 列和金额列的值，确保使用正确的列
          const snValue = (snColIndex !== null && snColIndex >= 0 && snColIndex < rowData.length) ? rowData[snColIndex] : "N/A";
          
          // 只有当金额列有值时才处理和显示
          if (amountValue !== null && amountValue !== undefined && amountValue !== "") {
            const cleanedAmount = cleanAmount(amountValue);
            totalAmount += cleanedAmount;
            amountValues.push(cleanedAmount);
            
            // 格式化原始值用于显示
            const rawValueStr = typeof amountValue === "string" ? `"${amountValue}"` : String(amountValue);
            
            // 追加调试日志到界面（同时显示 S/N 和金额列的值，便于验证）
            appendDebugLog(`Row ${excelRowNumber}: SN列[${snColIndex}]=${snValue}, 金额列[${amtColIndex}]=${rawValueStr} -> Clean=${cleanedAmount.toFixed(2)}`);
          } else {
            // 如果金额列为空，也记录一下
            appendDebugLog(`Row ${excelRowNumber}: SN列[${snColIndex}]=${snValue}, 金额列[${amtColIndex}]=空`);
          }
        } else {
          appendDebugLog(`Row ${excelRowNumber}: 警告 - 金额列索引 ${amtColIndex} 超出数据范围 (数据长度: ${rowData ? rowData.length : 0})`);
        }
      }
      
      // 输出汇总信息
      appendDebugLog(`--- 汇总 ---`);
      appendDebugLog(`总金额: ${totalAmount.toFixed(2)}`);
      appendDebugLog(`金额明细: [${amountValues.map(v => v.toFixed(2)).join(', ')}]`);
      
      // 创建新工作表
      const timestamp = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
      const newSheetName = `报表_${timestamp}`;
      let newSheet: Excel.Worksheet;
      
      try {
        newSheet = workbook.worksheets.add(newSheetName);
      } catch (error) {
        // 如果工作表已存在，使用带序号的名字
        let counter = 1;
        while (true) {
          try {
            newSheet = workbook.worksheets.add(`${newSheetName}_${counter}`);
            break;
          } catch (e) {
            counter++;
          }
        }
      }
      
      // 加载工作表名称，以便后续使用
      newSheet.load("name");
      
      // 从筛选面板获取当前应用的筛选条件（只记录用户选定的字段）
      // 格式：【字段名称】：（条件值）\n
      let currentFilterText = "无筛选条件";
      interface FilterDisplayItem {
        fieldName: string;
        values: string;
      }
      const filterDisplayItems: FilterDisplayItem[] = [];
      
      try {
        // 从筛选面板字段中获取有筛选的字段
        for (const field of filterPanelFields) {
          // 只记录有筛选的字段（选中的值少于全部值，且至少选中一个）
          if (field.selectedValues.size < field.allValues.length && field.selectedValues.size > 0) {
            const selectedValuesArray = Array.from(field.selectedValues);
            
            // 显示所有值，用分号分隔
            let displayValues: string;
            if (selectedValuesArray.length <= 5) {
              displayValues = selectedValuesArray.join("; ");
            } else {
              displayValues = selectedValuesArray.slice(0, 5).join("; ") + `... (共${selectedValuesArray.length}项)`;
            }
            
            filterDisplayItems.push({
              fieldName: field.headerText,
              values: displayValues
            });
            appendDebugLog(`筛选条件 - ${field.headerText}: ${selectedValuesArray.length}/${field.allValues.length}项`);
          }
        }
        
        if (filterDisplayItems.length > 0) {
          // 格式：【字段名称】：（条件值）\n
          currentFilterText = filterDisplayItems.map(item => 
            `【${item.fieldName}】：${item.values}`
          ).join("\n");
          appendDebugLog(`筛选条件: ${currentFilterText}`);
        } else {
          appendDebugLog("未检测到筛选条件（筛选面板无筛选）");
        }
      } catch (error) {
        console.warn("获取筛选条件失败:", error);
        appendDebugLog(`获取筛选条件失败: ${error.message}`);
      }
      
      // 计算 Dashboard 区域的行数（基础 4 行 + 筛选条件行数）
      const filterRowCount = filterDisplayItems.length > 0 ? filterDisplayItems.length : 1;
      const dashboardEndRow = 4 + filterRowCount + 1; // +1 for spacing row
      const dataStartRow = dashboardEndRow + 1;
      
      // 写入 Dashboard 区域 - 麦肯锡商务风格
      const dashboardRange = newSheet.getRange(`A1:E${dashboardEndRow}`);
      dashboardRange.format.verticalAlignment = Excel.VerticalAlignment.center;
      dashboardRange.format.horizontalAlignment = Excel.HorizontalAlignment.left;
      dashboardRange.format.fill.color = "#F5F5F5"; // 浅灰色背景
      dashboardRange.format.font.name = "Arial";
      dashboardRange.format.font.size = 11;
      dashboardRange.format.font.color = "#323130";
      
      // 设置 Dashboard 基础内容（前3行）
      // 标题使用 sheet 名称，便于同步修改
      const dashboardBaseData = [
        [newSheetName, "", "", "", ""],
        ["总条数", filteredCount.toString(), "", "源表名称", sourceSheetName],
        ["总金额", totalAmount.toFixed(2), "", "生成时间", new Date().toLocaleString("zh-CN")]
      ];
      const dashboardBaseRange = newSheet.getRange("A1:E3");
      dashboardBaseRange.values = dashboardBaseData;
      
      // 合并 Dashboard 标题单元格
      const titleCell = newSheet.getRange("A1:E1");
      titleCell.merge();
      titleCell.format.font.size = 14;
      titleCell.format.font.bold = true;
      titleCell.format.font.color = "#0078d4";
      titleCell.format.horizontalAlignment = Excel.HorizontalAlignment.center;
      
      // 设置 Dashboard 标签列（A列）的样式
      const labelRange = newSheet.getRange("A2:A3");
      labelRange.format.font.bold = true;
      labelRange.format.font.color = "#323130";
      
      // 设置 Dashboard 数据列的样式
      const dataRangeDashboard = newSheet.getRange("B2:B3");
      dataRangeDashboard.format.font.color = "#0078d4";
      dataRangeDashboard.format.font.bold = true;
      
      // 设置 Dashboard 右侧标签和数据的样式
      const rightLabelRange = newSheet.getRange("D2:D3");
      rightLabelRange.format.font.bold = true;
      const rightDataRange = newSheet.getRange("E2:E3");
      rightDataRange.format.font.color = "#323130";
      
      // 设置源表名称为超链接（点击返回原表标题行左起第一个单元格）
      const sourceSheetLinkCell = newSheet.getRange("E2");
      // 构建超链接地址：原表的标题行左起第一个单元格
      const configHeaderRow = currentMainReportConfig?.headerRow || 1;
      const linkAddress = `'${sourceSheetName}'!A${configHeaderRow}`;
      sourceSheetLinkCell.hyperlink = {
        address: "",
        documentReference: linkAddress,
        screenTip: `点击返回原表: ${sourceSheetName}`,
        textToDisplay: sourceSheetName
      };
      sourceSheetLinkCell.format.font.color = "#0078d4";
      sourceSheetLinkCell.format.font.underline = Excel.RangeUnderlineStyle.single;
      
      // 写入筛选条件区域（从第4行开始，每个筛选字段一行）
      // 格式：A列 = "筛选条件"，B列 = 字段名称（深蓝色），C列 = 条件值
      if (filterDisplayItems.length > 0) {
        for (let i = 0; i < filterDisplayItems.length; i++) {
          const rowNum = 4 + i;
          const item = filterDisplayItems[i];
          
          // 写入一整行数据（A-E列）
          const rowRange = newSheet.getRange(`A${rowNum}:E${rowNum}`);
          if (i === 0) {
            rowRange.values = [["筛选条件", `【${item.fieldName}】`, item.values, "", ""]];
          } else {
            rowRange.values = [["", `【${item.fieldName}】`, item.values, "", ""]];
          }
          
          // A 列样式
          const labelCell = newSheet.getRange(`A${rowNum}`);
          labelCell.format.font.bold = true;
          labelCell.format.font.color = "#323130";
          
          // B 列样式（字段名称，深蓝色）
          const fieldNameCell = newSheet.getRange(`B${rowNum}`);
          fieldNameCell.format.font.color = "#0d47a1"; // 深蓝色
          fieldNameCell.format.font.bold = true;
          
          // C 列样式（条件值）
          const valueCell = newSheet.getRange(`C${rowNum}`);
          valueCell.format.font.color = "#323130";
          valueCell.format.wrapText = true;
        }
      } else {
        // 无筛选条件
        const noFilterRow = newSheet.getRange("A4:E4");
        noFilterRow.values = [["筛选条件", "无筛选条件", "", "", ""]];
        
        const noFilterLabelCell = newSheet.getRange("A4");
        noFilterLabelCell.format.font.bold = true;
        noFilterLabelCell.format.font.color = "#323130";
        
        const noFilterValueCell = newSheet.getRange("B4");
        noFilterValueCell.format.font.color = "#a19f9d";
        noFilterValueCell.format.font.italic = true;
      }
      
      // 添加 Dashboard 边框
      dashboardRange.format.borders.getItem("EdgeTop").style = Excel.BorderLineStyle.continuous;
      dashboardRange.format.borders.getItem("EdgeTop").color = "#D1D1D1";
      dashboardRange.format.borders.getItem("EdgeBottom").style = Excel.BorderLineStyle.continuous;
      dashboardRange.format.borders.getItem("EdgeBottom").color = "#D1D1D1";
      dashboardRange.format.borders.getItem("EdgeLeft").style = Excel.BorderLineStyle.continuous;
      dashboardRange.format.borders.getItem("EdgeLeft").color = "#D1D1D1";
      dashboardRange.format.borders.getItem("EdgeRight").style = Excel.BorderLineStyle.continuous;
      dashboardRange.format.borders.getItem("EdgeRight").color = "#D1D1D1";
      dashboardRange.format.borders.getItem("InsideHorizontal").style = Excel.BorderLineStyle.continuous;
      dashboardRange.format.borders.getItem("InsideHorizontal").color = "#E1E1E1";
      dashboardRange.format.borders.getItem("InsideVertical").style = Excel.BorderLineStyle.continuous;
      dashboardRange.format.borders.getItem("InsideVertical").color = "#E1E1E1";
      
      // 先同步一下，确保前面的合并单元格操作完成
      await context.sync();
      
      // 写入数据（从 dataStartRow 行开始）- 麦肯锡商务风格
      if (filteredRows.length > 0) {
        const lastColumnName = getColumnName(columnCount - 1);
        
        // 调试：输出数据维度信息
        appendDebugLog(`准备写入数据...`);
        appendDebugLog(`目标范围: A${dataStartRow}:${lastColumnName}${dataStartRow + filteredRows.length - 1}`);
        appendDebugLog(`filteredRows.length = ${filteredRows.length}`);
        appendDebugLog(`columnCount = ${columnCount}`);
        
        // 验证每行的列数
        for (let i = 0; i < filteredRows.length; i++) {
          const rowLen = filteredRows[i] ? filteredRows[i].length : 0;
          if (rowLen !== columnCount) {
            appendDebugLog(`警告: 第 ${i} 行列数 (${rowLen}) 与 columnCount (${columnCount}) 不匹配`);
          }
        }
        
        // 计算目标范围的行列数
        const targetRowCount = filteredRows.length;
        const targetColCount = columnCount;
        appendDebugLog(`目标区域: ${targetRowCount} 行 x ${targetColCount} 列`);
        
        const dataRange = newSheet.getRange(`A${dataStartRow}:${lastColumnName}${dataStartRow + filteredRows.length - 1}`);
        
        try {
          dataRange.values = filteredRows;
        } catch (writeError) {
          appendDebugLog(`写入数据失败: ${writeError.message}`);
          // 尝试获取更多信息
          appendDebugLog(`第一行数据: ${JSON.stringify(filteredRows[0])}`);
          if (filteredRows.length > 1) {
            appendDebugLog(`第二行数据: ${JSON.stringify(filteredRows[1])}`);
          }
          throw writeError;
        }
        
        // 设置整体数据区域样式
        dataRange.format.font.name = "Arial";
        dataRange.format.font.size = 11;
        dataRange.format.font.color = "#323130";
        dataRange.format.verticalAlignment = Excel.VerticalAlignment.center;
        dataRange.format.horizontalAlignment = Excel.HorizontalAlignment.left;
        
        // 设置表头格式（dataStartRow行）
        const headerRange = newSheet.getRange(`A${dataStartRow}:${lastColumnName}${dataStartRow}`);
        headerRange.format.fill.color = "#0078d4"; // 麦肯锡蓝色背景
        headerRange.format.font.bold = true;
        headerRange.format.font.color = "#FFFFFF"; // 白色文字
        headerRange.format.font.size = 11;
        headerRange.format.font.name = "Arial";
        headerRange.format.horizontalAlignment = Excel.HorizontalAlignment.center;
        headerRange.format.verticalAlignment = Excel.VerticalAlignment.center;
        
        // 设置表头边框
        headerRange.format.borders.getItem("EdgeTop").style = Excel.BorderLineStyle.continuous;
        headerRange.format.borders.getItem("EdgeTop").color = "#005a9e";
        headerRange.format.borders.getItem("EdgeBottom").style = Excel.BorderLineStyle.continuous;
        headerRange.format.borders.getItem("EdgeBottom").color = "#005a9e";
        headerRange.format.borders.getItem("EdgeLeft").style = Excel.BorderLineStyle.continuous;
        headerRange.format.borders.getItem("EdgeLeft").color = "#005a9e";
        headerRange.format.borders.getItem("EdgeRight").style = Excel.BorderLineStyle.continuous;
        headerRange.format.borders.getItem("EdgeRight").color = "#005a9e";
        
        // 设置数据行样式（从dataStartRow+1行开始）
        if (filteredRows.length > 1) {
          const dataRowsRange = newSheet.getRange(`A${dataStartRow + 1}:${lastColumnName}${dataStartRow + filteredRows.length - 1}`);
          
          // 添加数据行边框
          dataRowsRange.format.borders.getItem("EdgeTop").style = Excel.BorderLineStyle.continuous;
          dataRowsRange.format.borders.getItem("EdgeTop").color = "#E1E1E1";
          dataRowsRange.format.borders.getItem("EdgeBottom").style = Excel.BorderLineStyle.continuous;
          dataRowsRange.format.borders.getItem("EdgeBottom").color = "#E1E1E1";
          dataRowsRange.format.borders.getItem("EdgeLeft").style = Excel.BorderLineStyle.continuous;
          dataRowsRange.format.borders.getItem("EdgeLeft").color = "#E1E1E1";
          dataRowsRange.format.borders.getItem("EdgeRight").style = Excel.BorderLineStyle.continuous;
          dataRowsRange.format.borders.getItem("EdgeRight").color = "#E1E1E1";
          dataRowsRange.format.borders.getItem("InsideHorizontal").style = Excel.BorderLineStyle.continuous;
          dataRowsRange.format.borders.getItem("InsideHorizontal").color = "#F0F0F0";
          dataRowsRange.format.borders.getItem("InsideVertical").style = Excel.BorderLineStyle.continuous;
          dataRowsRange.format.borders.getItem("InsideVertical").color = "#F0F0F0";
          
          // 交替行背景色（可选，更商务化）
          const dataRowStart = dataStartRow + 1;
          for (let i = dataRowStart; i <= dataStartRow + filteredRows.length - 1; i++) {
            const rowRange = newSheet.getRange(`A${i}:${lastColumnName}${i}`);
            if ((i - dataRowStart) % 2 === 0) {
              // 偶数行使用浅灰色背景
              rowRange.format.fill.color = "#FAFAFA";
            } else {
              // 奇数行使用白色背景
              rowRange.format.fill.color = "#FFFFFF";
            }
          }
        }
      }
      
      // 自动调整列宽
      const usedRangeNew = newSheet.getUsedRange();
      usedRangeNew.format.autofitColumns();
      
      // 设置行高
      const allRowsRange = newSheet.getUsedRange();
      allRowsRange.format.rowHeight = 20; // 统一行高
      
      // Dashboard 区域行高稍大
      const dashboardRows = newSheet.getRange(`1:${dashboardEndRow}`);
      dashboardRows.format.rowHeight = 22;
      
      // 激活新工作表
      newSheet.activate();
      
      await context.sync();
      
      // 添加报表到列表（此时 newSheet.name 已经加载）
      addReportToList(newSheet.name, filteredCount, totalAmount);
      
      // 保存当前筛选条件
      await saveCurrentFilterCondition();
      
      showMessage(`报表生成成功！共 ${filteredCount} 条记录，总金额: ${totalAmount.toFixed(2)}`);
    });
    
    // 更新筛选条件文本区域并自动复制到剪贴板
    updateFilterTextDisplay();
    await copyFilterTextToClipboard();
    
  } catch (error) {
    console.error("生成报表时出错:", error);
    showMessage(`生成报表失败: ${error.message}`, true);
  }
}

// 加载主报表配置
async function loadMainReportConfig(): Promise<void> {
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      
      // 尝试加载数据区域
      try {
        const dataRangeItem = workbook.names.getItem(CFG_DATA_RANGE_NAME);
        dataRangeItem.load("formula");
        await context.sync();
        
        const rangeMatch = dataRangeItem.formula.match(/=([^!]+)!([^$]+)/);
        if (rangeMatch) {
          if (!currentMainReportConfig) {
            currentMainReportConfig = {
              dataRange: dataRangeItem.formula.replace(/^=/, ""),
              headerRow: 0,
              snColumn: 0,
              amtColumn: 0,
              sheetName: rangeMatch[1]
            };
          } else {
            currentMainReportConfig.dataRange = dataRangeItem.formula.replace(/^=/, "");
            currentMainReportConfig.sheetName = rangeMatch[1];
          }
        }
      } catch (error) {
        // 数据区域未配置
      }
      
      // 尝试加载标题行
      try {
        const headerRowItem = workbook.names.getItem(CFG_HEADER_ROW_NAME);
        headerRowItem.load("formula");
        await context.sync();
        
        const rowMatch = headerRowItem.formula.match(/:(\d+):/);
        if (rowMatch && currentMainReportConfig) {
          currentMainReportConfig.headerRow = parseInt(rowMatch[1]);
        }
      } catch (error) {
        // 标题行未配置
      }
      
      // 尝试加载 S/N 列
      try {
        const snItem = workbook.names.getItem(CFG_SN_COL_NAME);
        const snRange = snItem.getRange();
        snRange.load("columnIndex");
        await context.sync();
        
        if (currentMainReportConfig) {
          currentMainReportConfig.snColumn = snRange.columnIndex;
        }
      } catch (error) {
        // S/N 列未配置
      }
      
      // 尝试加载金额列
      try {
        const amtItem = workbook.names.getItem(CFG_AMT_COL_NAME);
        const amtRange = amtItem.getRange();
        amtRange.load("columnIndex");
        await context.sync();
        
        if (currentMainReportConfig) {
          currentMainReportConfig.amtColumn = amtRange.columnIndex;
        }
      } catch (error) {
        // 金额列未配置
      }
    });
  } catch (error) {
    console.error("加载主报表配置时出错:", error);
  }
}

// 初始化
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    // 绑定按钮事件
    document.getElementById("btn-set-data-range")?.addEventListener("click", () => {
      setDataRange();
    });
    
    document.getElementById("btn-set-header-row")?.addEventListener("click", () => {
      setHeaderRow();
    });
    
    document.getElementById("btn-set-sn")?.addEventListener("click", () => {
      setSnColumnWithMergeCheck();
    });
    
    document.getElementById("btn-set-amt")?.addEventListener("click", () => {
      setColumnConfig(CFG_AMT_COL_NAME, "金额列");
    });
    
    document.getElementById("btn-enable-filter")?.addEventListener("click", () => {
      enableFilter();
    });
    
    document.getElementById("btn-generate")?.addEventListener("click", () => {
      generateReport();
    });
    
    // 筛选面板按钮事件
    document.getElementById("btn-apply-filter")?.addEventListener("click", () => {
      applyFilterFromPanel();
    });
    
    document.getElementById("btn-clear-filter")?.addEventListener("click", () => {
      clearFilterFromPanel();
    });
    
    document.getElementById("btn-save-filter")?.addEventListener("click", () => {
      saveFilterFromPanel();
    });
    
    // 重置筛选按钮事件
    document.getElementById("btn-reset-filter")?.addEventListener("click", () => {
      resetFilterAndGoToMainReport();
    });
    
    // 点击主报表区域文字跳转到主报表
    document.getElementById("data-range-info")?.addEventListener("click", () => {
      goToMainReport();
    });
    
    // 绑定筛选面板事件委托（只绑定一次）
    bindFilterPanelEvents();
    
    // 绑定筛选条件文本区域点击事件
    bindFilterTextEvents();
    
    // 初始化显示
    loadMainReportConfig().then(() => {
      updateMainReportConfigDisplay();
      updateReportsTable();
      updateFilterConditionsList();
    });
    
    console.log("Excel Add-in 初始化完成");
  }
});

