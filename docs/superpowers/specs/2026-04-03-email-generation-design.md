# 邮件生成与数据格式统一设计

## 1. 概述

本次改造包含两个相关问题的修复：

1. **数据格式统一** - 统一邮件附件 Excel、邮件预览 HTML 的表头与列格式
2. **Excel 样式支持** - 从 xlsx 库迁移到 ExcelJS，支持单元格样式（合并居中等）

> **注**：邮件生成最终保留手动 MIME 拼接（未使用 eml-generator 库），原因是 eml-generator 对字符串数据的 Base64 处理存在 bug，会将附件内容变为无效文件。

---

## 2. 数据格式对比

| 数据位置 | 表头格式 | 列内容 | 行数 |
|----------|----------|--------|------|
| 页面表格 | 单层 | 完整列 | 全部 |
| Excel 下载 | 双层合并 | 完整列 | 全部 |
| 邮件附件 Excel | 双层合并 | 完整列 | 全部 |
| 邮件预览 HTML | 分组标签 | 完整列 | 前20行 + 备注 |

**统一列顺序**（全局定义）：
```
物料 → 物料描述 → 基本计量单位 → 分母i → 可选单位i → = → 计数器i → 基本计量单位i
```

---

## 3. 全局表头定义

### 3.1 导出函数（复用）

```javascript
// getTableColumns(maxUnits)      → 单层表头列定义（页面表格用）
// getHeaderRow1(maxUnits)         → 双层表头第1行（分组大标题）
// getHeaderRow2(maxUnits)         → 双层表头第2行（具体列名）
// getColumnProps(maxUnits)        → 列 prop 数组
```

### 3.2 列顺序统一

`getTableColumns` / `getHeaderRow2` / `getColumnProps` 三者顺序完全一致，维护成本降低。

---

## 4. Excel 库迁移（xlsx → ExcelJS）

### 4.1 原因

xlsx 库无法设置单元格样式（合并单元格的居中、边框等）。切换到 ExcelJS 以支持完整样式。

### 4.2 API 对照

| 功能 | xlsx | ExcelJS |
|------|------|---------|
| 读取 | `XLSX.read()` 同步 | `workbook.xlsx.load()` 异步 |
| Sheet→数组 | `sheet_to_json({header:1})` | `eachRow()` 手动转 |
| 数组→Sheet | `aoa_to_sheet()` | `addRow()` |
| 合并单元格 | `ws['!merges']=` | `worksheet.mergeCells()` |
| 样式 | `ws[ref].s=` | `cell.alignment=` / `cell.border=` |
| 列宽 | `ws['!cols']=` | `worksheet.getColumn().width` |
| 下载 | `writeFile()` | `writeBuffer()` → Blob |
| Base64 | `write(...,{type:'base64'})` | `writeBuffer()` → bufferToBase64() |

### 4.3 共享工作簿构建

`buildExcelWorkbook(mergedData, maxUnits)` 是导出 Excel 和邮件附件共用的工作簿构建函数，保证两者格式完全一致：

- 双层表头合并（第一行）
- 前3列垂直合并（A1:A2, B1:B2, C1:C2）
- 全局居中 + 全边框样式
- 物料描述列（B）宽度 70
- 第一行高度 30

---

## 5. 邮件生成（手动 MIME）

### 5.1 Base64 转换

由于放弃了 eml-generator，邮件附件 Base64 采用 `bufferToBase64()` 函数（使用 FileReader API）生成，避免大文件 stack overflow：

```javascript
function bufferToBase64(buffer) {
  return new Promise((resolve) => {
    const blob = new Blob([buffer], { type: 'application/octet-stream' })
    const reader = new FileReader()
    reader.onloadend = () => resolve(reader.result.split(',')[1])
    reader.readAsDataURL(blob)
  })
}
```

### 5.2 EML 结构

```
MIME-Version: 1.0
Content-Type: multipart/mixed; boundary="..."
From: sap-system@company.com
To: <收件人>
Subject: <主题>
Date: <UTC时间>

--boundary
Content-Type: text/html; charset=utf-8
[HTML 预览正文]

--boundary
Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
Content-Transfer-Encoding: base64
Content-Disposition: attachment; filename="material_data.xlsx"
[Base64 编码的 Excel 数据]

--boundary--
```

---

## 6. 依赖变更

| 操作 | 包 |
|------|-----|
| 新增 | `exceljs` ^4.4.0 |
| 移除 | `xlsx` |
| 移除 | `eml-generator`（未采用） |

---

## 7. 改动文件清单

| 文件 | 改动内容 |
|------|----------|
| `excelProcessor.js` | 1. xlsx → ExcelJS 全面迁移<br>2. 新增 `buildExcelWorkbook` 共享函数<br>3. `getHeaderRow1/2` 导出复用<br>4. `bufferToBase64` 替代 writeBase64<br>5. 列顺序全局统一 |
| `MaterialMergeView.vue` | `handleExportExcel` 和 `confirmExportEmail` 改为 async/await |
| `package.json` | exceljs 新增，xlsx 移除 |

---

## 8. 测试要点

1. **Excel 读取** - 上传两个 Excel 文件，验证行数、列数、空单元格处理
2. **Excel 导出** - 验证双层表头、合并居中、列宽、边框样式，文件可用 Excel/WPS 打开
3. **邮件预览 HTML** - 显示完整列（前20行），底部有"此为数据预览，详见附件"备注
4. **邮件附件** - EML 中附件 Base64 正确，下载后扩展名为 .xlsx 且可正常打开
5. **回归对比** - 迁移前后数据输出一致
