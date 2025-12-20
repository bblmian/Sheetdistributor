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

// 显示消息到消息区域
function showMessage(message: string, isError: boolean = false): void {
  const messageArea = document.getElementById("message-area");
  if (messageArea) {
    messageArea.textContent = message;
    messageArea.className = isError ? "message-area error" : "message-area success";
  }
  console.log(message);
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
      selection.load("address");
      
      await context.sync();
      
      // 获取选中范围的整列
      const selectedColumn = selection.getEntireColumn();
      selectedColumn.load("columnIndex");
      
      await context.sync();
      
      const columnIndex = selectedColumn.columnIndex;
      const columnNameStr = getColumnName(columnIndex);
      
      // 创建或更新 NamedItem
      const namedItems = context.workbook.names;
      
      // 先尝试删除已存在的（如果存在）
      try {
        const existingItem = namedItems.getItem(columnName);
        existingItem.delete();
        await context.sync();
      } catch (error) {
        // 如果不存在，忽略错误
      }
      
      // 创建新的 NamedItem，引用到选中的列
      // 使用绝对引用格式：=Sheet1!$A:$A
      const columnAddress = `=${sheet.name}!$${columnNameStr}:$${columnNameStr}`;
      namedItems.add(columnName, columnAddress);
      
      await context.sync();
      
      // 验证：读取刚创建的 NamedItem，确认列索引正确
      const verifyItem = namedItems.getItem(columnName);
      verifyItem.load("name, formula");
      const verifyRange = verifyItem.getRange();
      verifyRange.load("columnIndex, address");
      await context.sync();
      
      const verifyColIndex = verifyRange.columnIndex;
      if (verifyColIndex !== columnIndex) {
        console.warn(`警告: NamedItem 列索引不匹配！期望: ${columnIndex}, 实际: ${verifyColIndex}`);
      }
      
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
      
      showMessage(`成功设置${displayName}为第 ${columnIndex + 1} 列 (${columnNameStr})，验证列索引: ${verifyColIndex + 1}`);
      
      // 更新配置显示
      updateMainReportConfigDisplay();
    });
  } catch (error) {
    console.error("设置列配置时出错:", error);
    showMessage(`设置${displayName}失败: ${error.message}`, true);
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
      sheet.name = newName;
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
  
  container.innerHTML = `
    <div class="current-filter-title">当前应用的筛选条件：</div>
    <div class="current-filter-items">
      ${activeFilters.map(filter => `
        <div class="current-filter-item">
          <span class="filter-field">${filter.headerText}</span>
          <span class="filter-separator">：</span>
          <span class="filter-value">${filter.filterValues.join(", ")}</span>
        </div>
      `).join("")}
    </div>
  `;
}

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
      
      // 只显示前5个值，避免太长
      let displayValues: string[];
      if (selectedValuesArray.length <= 5) {
        displayValues = selectedValuesArray;
      } else {
        displayValues = selectedValuesArray.slice(0, 5);
        displayValues.push(`... (共${selectedValuesArray.length}项)`);
      }
      
      columnFilterSettings.push({
        columnIndex: field.columnIndex,
        columnName: getColumnName(field.columnIndex),
        headerText: field.headerText,
        filterType: "values",
        filterValues: displayValues,
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
        <div class="filter-select-actions">
          <button class="filter-select-action" data-field-index="${activeFilterFieldIndex}" data-action="all">全选</button>
          <button class="filter-select-action" data-field-index="${activeFilterFieldIndex}" data-action="none">全不选</button>
          <button class="filter-select-action" data-field-index="${activeFilterFieldIndex}" data-action="invert">反选</button>
        </div>
      </div>
    `;
  }
  
  panel.innerHTML = tagsHtml + dropdownHtml;
  
  // 绑定事件
  bindFilterPanelEvents();
}

// HTML 转义函数
function escapeHtml(text: string): string {
  const div = document.createElement("div");
  div.textContent = text;
  return div.innerHTML;
}

// 绑定筛选面板事件
function bindFilterPanelEvents(): void {
  const panel = document.getElementById("filter-panel");
  if (!panel) return;
  
  // 标签点击事件 - 切换激活的字段
  const tags = panel.querySelectorAll(".filter-tag");
  tags.forEach((tag) => {
    tag.addEventListener("click", (e) => {
      const target = e.currentTarget as HTMLElement;
      const fieldIndex = parseInt(target.getAttribute("data-field-index") || "-1", 10);
      
      if (fieldIndex === activeFilterFieldIndex) {
        // 如果点击的是当前激活的标签，则收起
        activeFilterFieldIndex = -1;
      } else {
        // 否则切换到新的字段
        activeFilterFieldIndex = fieldIndex;
      }
      
      // 重新渲染面板
      renderFilterPanel();
    });
  });
  
  // 关闭下拉框按钮
  const closeBtn = document.getElementById("btn-close-dropdown");
  if (closeBtn) {
    closeBtn.addEventListener("click", () => {
      activeFilterFieldIndex = -1;
      renderFilterPanel();
    });
  }
  
  // 搜索框事件
  const searchInputs = panel.querySelectorAll(".filter-field-search");
  searchInputs.forEach((input) => {
    input.addEventListener("input", (e) => {
      const target = e.target as HTMLInputElement;
      const fieldIndex = parseInt(target.getAttribute("data-field-index") || "0", 10);
      filterSearchOptions(fieldIndex, target.value);
    });
  });
  
  // 复选框点击事件
  panel.addEventListener("change", (e) => {
    const target = e.target as HTMLInputElement;
    if (target.type === "checkbox") {
      const fieldIndex = parseInt(target.getAttribute("data-field-index") || "0", 10);
      const value = target.getAttribute("data-value") || "";
      
      if (fieldIndex >= 0 && fieldIndex < filterPanelFields.length) {
        const field = filterPanelFields[fieldIndex];
        if (target.checked) {
          field.selectedValues.add(value);
        } else {
          field.selectedValues.delete(value);
        }
        
        // 更新选项的 selected 样式
        const optionDiv = target.closest(".filter-option");
        if (optionDiv) {
          optionDiv.classList.toggle("selected", target.checked);
        }
        
        // 更新标签上的计数
        updateTagCount(fieldIndex);
      }
    }
  });
  
  // 全选/全不选/反选按钮事件
  panel.addEventListener("click", (e) => {
    const target = e.target as HTMLElement;
    if (target.classList.contains("filter-select-action")) {
      const fieldIndex = parseInt(target.getAttribute("data-field-index") || "0", 10);
      const action = target.getAttribute("data-action");
      
      if (fieldIndex >= 0 && fieldIndex < filterPanelFields.length) {
        const field = filterPanelFields[fieldIndex];
        
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
          field.selectedValues = newSelected;
        }
        
        // 重新渲染下拉框中的选项
        updateFieldOptions(fieldIndex);
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
  if (!optionsContainer) return;
  
  const options = optionsContainer.querySelectorAll(".filter-option");
  const searchLower = searchText.toLowerCase();
  
  options.forEach((option) => {
    const value = option.getAttribute("data-value") || "";
    const matches = value.toLowerCase().includes(searchLower);
    option.classList.toggle("hidden", !matches);
  });
}

// 更新字段的选项（用于全选/反选后）
function updateFieldOptions(fieldIndex: number): void {
  const field = filterPanelFields[fieldIndex];
  const optionsContainer = document.getElementById(`filter-field-${fieldIndex}-options`);
  if (!optionsContainer || !field) return;
  
  const options = optionsContainer.querySelectorAll(".filter-option");
  options.forEach((option) => {
    const value = option.getAttribute("data-value") || "";
    const isSelected = field.selectedValues.has(value);
    const checkbox = option.querySelector("input[type='checkbox']") as HTMLInputElement;
    if (checkbox) {
      checkbox.checked = isSelected;
    }
    option.classList.toggle("selected", isSelected);
  });
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
      
      // 获取使用范围（包括起始行信息）
      const usedRange = targetSheet.getUsedRange();
      usedRange.load("values, rowCount, rowIndex");
      await context.sync();
      
      if (!usedRange.values || usedRange.values.length === 0) {
        showMessage("工作表没有数据", true);
        return;
      }
      
      // usedRange.rowIndex 是 0-based，第一行是 0
      const startRow = usedRange.rowIndex + 1; // 转换为 1-based Excel 行号
      const endRow = startRow + usedRange.rowCount - 1;
      
      // 构建筛选条件 - 只记录有筛选的字段
      const filterEntries: Array<{ colIdx: number; allowedValues: Set<string>; headerText: string }> = [];
      
      for (const field of filterPanelFields) {
        // 如果该列有筛选（选中的值少于全部值，且至少选中一个）
        if (field.selectedValues.size < field.allValues.length && field.selectedValues.size > 0) {
          filterEntries.push({
            colIdx: field.columnIndex,
            allowedValues: field.selectedValues,
            headerText: field.headerText
          });
        }
      }
      
      // 首先显示所有数据行（从标题行的下一行开始）
      const dataStartRow = startRow + 1; // 跳过标题行
      for (let excelRow = dataStartRow; excelRow <= endRow; excelRow++) {
        const row = targetSheet.getRange(`${excelRow}:${excelRow}`);
        row.rowHidden = false;
      }
      await context.sync();
      
      // 如果没有筛选条件，直接返回
      if (filterEntries.length === 0) {
        updateCurrentFilterDisplay([]);
        showMessage("已清除所有筛选，显示全部数据");
        return;
      }
      
      // 遍历每一行数据，应用筛选
      let hiddenCount = 0;
      let visibleCount = 0;
      
      // usedRange.values[0] 是标题行，从 [1] 开始是数据行
      for (let valueIdx = 1; valueIdx < usedRange.values.length; valueIdx++) {
        const rowData = usedRange.values[valueIdx];
        const excelRowNum = startRow + valueIdx; // 对应的 Excel 行号
        let shouldHide = false;
        
        // 检查该行是否符合所有筛选条件
        for (const filter of filterEntries) {
          if (filter.colIdx < rowData.length) {
            const cellValue = rowData[filter.colIdx];
            const cellStr = cellValue !== null && cellValue !== undefined ? String(cellValue) : "";
            
            // 如果该单元格的值不在选中的值集合中，则隐藏该行
            if (!filter.allowedValues.has(cellStr)) {
              shouldHide = true;
              break;
            }
          }
        }
        
        if (shouldHide) {
          const row = targetSheet.getRange(`${excelRowNum}:${excelRowNum}`);
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
      usedRange.load("rowCount");
      await context.sync();
      
      // 显示所有行
      for (let i = 2; i <= usedRange.rowCount; i++) {
        const row = targetSheet.getRange(`${i}:${i}`);
        row.rowHidden = false;
      }
      await context.sync();
      
      updateCurrentFilterDisplay([]);
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
  
  try {
    await Excel.run(async (context) => {
      const workbook = context.workbook;
      const sheet = workbook.worksheets.getItem(currentMainReportConfig!.sheetName);
      sheet.activate();
      
      // 获取使用范围
      const usedRange = sheet.getUsedRange();
      usedRange.load("rowCount, rowIndex");
      await context.sync();
      
      // 显示所有行
      const startRow = usedRange.rowIndex + 1;
      const endRow = startRow + usedRange.rowCount - 1;
      
      for (let excelRow = startRow + 1; excelRow <= endRow; excelRow++) {
        const row = sheet.getRange(`${excelRow}:${excelRow}`);
        row.rowHidden = false;
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
      const sourceSheet = workbook.worksheets.getActiveWorksheet();
      sourceSheet.load("name");
      
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
      
      usedRange.load("rowCount, columnCount, address, values");
      
      // 检查每一行是否可见（未被筛选隐藏）
      const rows: Excel.Range[] = [];
      // 先获取行数，但需要先 sync
      await context.sync();
      
      if (usedRange.rowCount === 0 || usedRange.columnCount === 0) {
        showMessage("错误: 当前工作表没有数据", true);
        return;
      }
      
      const rowCount = usedRange.rowCount;
      const columnCount = usedRange.columnCount;
      
      // 检查每一行是否可见（未被筛选隐藏）
      for (let i = 1; i <= rowCount; i++) {
        const row = sourceSheet.getRange(`${i}:${i}`);
        row.load("rowHidden");
        rows.push(row);
      }
      
      await context.sync();
      
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
      
      // 在内存中处理数据
      const filteredRows: any[][] = [];
      
      // 第一行通常是表头，始终保留（不累加金额）
      if (usedRange.values.length > 0 && usedRange.values[0]) {
        filteredRows.push(usedRange.values[0]);
      }
      
      // 从第二行开始，只处理可见的行（rowHidden === false）
      // rows 数组索引从 0 开始，对应第 1 行；usedRange.values 索引也从 0 开始
      let visibleRowCount = 0;
      for (let i = 1; i < rowCount && i < rows.length && i < usedRange.values.length; i++) {
        const row = rows[i];
        
        // 只处理可见的行（未被筛选隐藏的行）
        if (row && !row.rowHidden) {
          const rowData = usedRange.values[i];
          if (rowData && Array.isArray(rowData)) {
            filteredRows.push(rowData);
            visibleRowCount++;
          }
        }
      }
      
      const filteredCount = filteredRows.length - 1; // 减去表头
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
      let currentFilterText = "无筛选条件";
      try {
        const filterItems: string[] = [];
        
        // 从筛选面板字段中获取有筛选的字段
        for (const field of filterPanelFields) {
          // 只记录有筛选的字段（选中的值少于全部值，且至少选中一个）
          if (field.selectedValues.size < field.allValues.length && field.selectedValues.size > 0) {
            const selectedValuesArray = Array.from(field.selectedValues);
            
            // 只显示前5个值，避免太长
            let displayValues: string;
            if (selectedValuesArray.length <= 5) {
              displayValues = selectedValuesArray.join(", ");
            } else {
              displayValues = selectedValuesArray.slice(0, 5).join(", ") + `... (共${selectedValuesArray.length}项)`;
            }
            
            filterItems.push(`${field.headerText}: ${displayValues}`);
            appendDebugLog(`筛选条件 - ${field.headerText}: ${selectedValuesArray.length}/${field.allValues.length}项`);
          }
        }
        
        if (filterItems.length > 0) {
          currentFilterText = filterItems.join("; ");
          appendDebugLog(`筛选条件: ${currentFilterText}`);
        } else {
          // 如果筛选面板没有筛选条件，检查是否有行被隐藏
          let hiddenCount = 0;
          for (let i = 1; i < rows.length; i++) {
            if (rows[i] && rows[i].rowHidden) {
              hiddenCount++;
            }
          }
          if (hiddenCount > 0) {
            currentFilterText = `已筛选 (隐藏${hiddenCount}行)`;
            appendDebugLog(`有${hiddenCount}行被隐藏`);
          } else {
            appendDebugLog("未检测到筛选条件");
          }
        }
      } catch (error) {
        console.warn("获取筛选条件失败:", error);
        appendDebugLog(`获取筛选条件失败: ${error.message}`);
      }
      
      // 写入 Dashboard 区域 (Row 1-6) - 麦肯锡商务风格（增加一行显示筛选条件）
      const dashboardRange = newSheet.getRange("A1:E6");
      dashboardRange.format.verticalAlignment = Excel.VerticalAlignment.center;
      dashboardRange.format.horizontalAlignment = Excel.HorizontalAlignment.left;
      dashboardRange.format.fill.color = "#F5F5F5"; // 浅灰色背景
      dashboardRange.format.font.name = "Arial";
      dashboardRange.format.font.size = 11;
      dashboardRange.format.font.color = "#323130";
      
      // 设置 Dashboard 内容
      const dashboardData = [
        ["报表 Dashboard", "", "", "", ""],
        ["总条数", filteredCount.toString(), "", "源表名称", sourceSheetName],
        ["总金额", totalAmount.toFixed(2), "", "生成时间", new Date().toLocaleString("zh-CN")],
        ["筛选条件", currentFilterText, "", "", ""],
        ["", "", "", "", ""],
        ["", "", "", "", ""]
      ];
      
      dashboardRange.values = dashboardData;
      
      // 合并 Dashboard 标题单元格
      const titleCell = newSheet.getRange("A1:E1");
      titleCell.merge();
      titleCell.format.font.size = 14;
      titleCell.format.font.bold = true;
      titleCell.format.font.color = "#0078d4";
      titleCell.format.horizontalAlignment = Excel.HorizontalAlignment.center;
      
      // 设置 Dashboard 标签列（A列）的样式
      const labelRange = newSheet.getRange("A2:A4");
      labelRange.format.font.bold = true;
      labelRange.format.font.color = "#323130";
      
      // 设置 Dashboard 数据列的样式
      const dataRange = newSheet.getRange("B2:B3");
      dataRange.format.font.color = "#0078d4";
      dataRange.format.font.bold = true;
      
      // 设置筛选条件行的样式
      const filterLabelCell = newSheet.getRange("A4");
      filterLabelCell.format.font.bold = true;
      filterLabelCell.format.font.color = "#323130";
      const filterValueCell = newSheet.getRange("B4:E4");
      filterValueCell.merge();
      filterValueCell.format.font.color = "#605e5c";
      filterValueCell.format.wrapText = true;
      
      // 设置 Dashboard 右侧标签和数据的样式
      const rightLabelRange = newSheet.getRange("D2:D3");
      rightLabelRange.format.font.bold = true;
      const rightDataRange = newSheet.getRange("E2:E3");
      rightDataRange.format.font.color = "#323130";
      
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
      
      // 写入数据（从第 6 行开始）- 麦肯锡商务风格
      if (filteredRows.length > 0) {
        const lastColumnName = getColumnName(columnCount - 1);
        const dataRange = newSheet.getRange(`A6:${lastColumnName}${6 + filteredRows.length - 1}`);
        dataRange.values = filteredRows;
        
        // 设置整体数据区域样式
        dataRange.format.font.name = "Arial";
        dataRange.format.font.size = 11;
        dataRange.format.font.color = "#323130";
        dataRange.format.verticalAlignment = Excel.VerticalAlignment.center;
        dataRange.format.horizontalAlignment = Excel.HorizontalAlignment.left;
        
        // 设置表头格式（第6行）
        const headerRange = newSheet.getRange(`A6:${lastColumnName}6`);
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
        
        // 设置数据行样式（从第7行开始）
        if (filteredRows.length > 1) {
          const dataRowsRange = newSheet.getRange(`A7:${lastColumnName}${6 + filteredRows.length - 1}`);
          
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
          for (let i = 7; i <= 6 + filteredRows.length; i++) {
            const rowRange = newSheet.getRange(`A${i}:${lastColumnName}${i}`);
            if ((i - 7) % 2 === 0) {
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
      const dashboardRows = newSheet.getRange("1:6");
      dashboardRows.format.rowHeight = 25;
      
      // 筛选条件行可能需要更大的高度（如果内容较多）
      const filterRow = newSheet.getRange("4:4");
      filterRow.format.rowHeight = 30;
      
      // 激活新工作表
      newSheet.activate();
      
      await context.sync();
      
      // 添加报表到列表（此时 newSheet.name 已经加载）
      addReportToList(newSheet.name, filteredCount, totalAmount);
      
      // 保存当前筛选条件
      await saveCurrentFilterCondition();
      
      showMessage(`报表生成成功！共 ${filteredCount} 条记录，总金额: ${totalAmount.toFixed(2)}`);
    });
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
      setColumnConfig(CFG_SN_COL_NAME, "S/N 列");
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
    
    // 初始化显示
    loadMainReportConfig().then(() => {
      updateMainReportConfigDisplay();
      updateReportsTable();
      updateFilterConditionsList();
    });
    
    console.log("Excel Add-in 初始化完成");
  }
});
