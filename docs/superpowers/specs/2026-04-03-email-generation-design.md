# 邮件生成与数据格式统一设计

## 1. 概述

本次改造包含两个独立但相关的问题修复：

1. **邮件生成库替换** - 将手动拼接 EML 格式替换为 Mailcomposer 库
2. **数据格式统一** - 统一邮件附件 Excel、邮件预览 HTML 的表头与列格式

---

## 2. 问题 #1：邮件生成库替换

### 现状

`MaterialMergeView.vue` 中的 `confirmExportEmail` 函数手动拼接 RFC 822 MIME 格式：
- 手动构造 `From`, `To`, `Subject`, `Date`, `MIME-Version`, `Content-Type` 等头部
- 手动处理 Base64 编码的附件
- 缺少对中文文件名的 RFC 2231/2047 编码支持
- 代码冗长，难以维护

### 解决方案

安装并使用 **eml-generator** 库（零依赖，支持浏览器和 Node.js）：

```bash
npm i eml-generator
```

### 改动文件

#### `excelProcessor.js` - 新增函数

**`generateAttachmentExcelBuffer(mergedData, maxUnits)`** - 生成邮件附件 Excel 的 Buffer（双层表头格式）

```javascript
/**
 * 生成邮件附件 Excel Buffer（双层表头）
 * @param {Array} mergedData - 合并后数据
 * @param {number} maxUnits - 最大单位数
 * @returns {Buffer} Excel 文件 Buffer
 */
export function generateAttachmentExcelBuffer(mergedData, maxUnits) {
  const props = ['物料', '物料描述', '基本计量单位']
  for (let i = 1; i <= maxUnits; i++) {
    props.push(`分母${i}`, `基本计量单位${i}`, `可选单位${i}`, `等于${i}`, `计数器${i}`)
  }

  const headerRow1 = ['物料', '物料描述', '基本计量单位']
  const headerRow2 = ['物料', '物料描述', '基本计量单位']

  for (let i = 1; i <= maxUnits; i++) {
    headerRow1.push(`单位转换${i}`, '', '', '', '')
    headerRow2.push(`分母${i}`, `基本计量单位${i}`, `可选单位${i}`, `等于${i}`, `计数器${i}`)
  }

  const dataRows = mergedData.map(row => props.map(prop => row[prop] || ''))
  const allData = [headerRow1, headerRow2, ...dataRows]

  const ws = XLSX.utils.aoa_to_sheet(allData)

  const merges = []
  let mergeCol = 3
  for (let i = 1; i <= maxUnits; i++) {
    merges.push({ s: { r: 0, c: mergeCol }, e: { r: 0, c: mergeCol + 4 } })
    mergeCol += 5
  }
  ws['!merges'] = merges
  ws['!cols'] = props.map((_, idx) => idx < 3 ? { wch: 15 } : { wch: 12 })

  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, 'material_data')

  return XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' })
}
```

**`generateEmlContent(mergedData, maxUnits, emailOptions)`** - 生成 .eml 格式字符串

```javascript
import { eml } from 'eml-generator'

/**
 * 生成 EML 格式字符串
 * @param {Array} mergedData - 合并后数据
 * @param {number} maxUnits - 最大单位数
 * @param {Object} emailOptions - 邮件选项 { to, from, subject }
 * @returns {string} EML 格式字符串
 */
export function generateEmlContent(mergedData, maxUnits, emailOptions) {
  // 生成附件 Excel Buffer（双层表头）
  const excelBuffer = generateAttachmentExcelBuffer(mergedData, maxUnits)

  // 生成 HTML 预览（完整列，前20行，带备注）
  const previewHtml = generateEmailPreviewHtml(mergedData, maxUnits)

  const emailContent = eml({
    from: emailOptions.from || 'sap-system@company.com',
    to: emailOptions.to,
    subject: emailOptions.subject || 'SAP物料单位数据',
    text: `SAP物料单位数据汇总（共 ${mergedData.length} 条记录）`,
    html: previewHtml,
    attachments: [
      {
        filename: 'material_data.xlsx',
        data: excelBuffer,
        contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      }
    ]
  })

  return emailContent
}
```

#### `MaterialMergeView.vue` - 改造 `confirmExportEmail`

```javascript
import { generateEmlContent } from '../utils/excelProcessor'

function confirmExportEmail() {
  if (!emailForm.value.to) {
    ElMessage.warning('请输入收件人邮箱')
    return
  }

  try {
    const emlContent = generateEmlContent(
      mergedData.value,
      maxUnits.value,
      {
        to: emailForm.value.to,
        subject: emailForm.value.subject || 'SAP物料单位数据',
        from: 'sap-system@company.com'
      }
    )

    // 下载 EML 文件
    const blob = new Blob([emlContent], { type: 'message/rfc822' })
    const url = URL.createObjectURL(blob)
    const a = document.createElement('a')
    a.href = url
    a.download = 'SAP物料数据.eml'
    a.click()
    URL.revokeObjectURL(url)

    emailDialogVisible.value = false
    ElMessage.success('邮件文件已生成并下载')
  } catch (err) {
    console.error(err)
    ElMessage.error('生成邮件失败')
  }
}
```

---

## 3. 问题 #2：数据格式统一

### 现状

| 数据位置 | 表头格式 | 列内容 |
|----------|----------|--------|
| 页面表格 | 单层 | 完整列 |
| Excel 下载 | 双层合并 | 完整列 |
| 邮件附件 Excel | 单层 | 完整列（bug） |
| 邮件预览 HTML | 单层 | **只有3列（缺单位转换列）** |

### 解决方案

#### 3.1 邮件附件 Excel - 统一为双层表头

复用在 `exportToExcel` 中已有的双层表头逻辑，替换 `generateAttachmentBase64` 函数为 `generateAttachmentExcelBuffer`（见上文）。

#### 3.2 邮件预览 HTML - 完整列 + 预览备注

**`excelProcessor.js` - 新增 `generateEmailPreviewHtml`**

```javascript
/**
 * 生成邮件预览 HTML（完整列，前20行，带备注）
 * @param {Array} data - 数据（通常传前20条）
 * @param {number} maxUnits - 最大单位数
 * @returns {string} HTML 字符串
 */
export function generateEmailPreviewHtml(data, maxUnits) {
  if (!data || data.length === 0) return ''

  // 构建完整列 props（与 Excel 一致）
  const baseProps = ['物料', '物料描述', '基本计量单位']
  const unitProps = []
  for (let i = 1; i <= maxUnits; i++) {
    unitProps.push(`分母${i}`, `基本计量单位${i}`, `可选单位${i}`, `等于${i}`, `计数器${i}`)
  }
  const allProps = [...baseProps, ...unitProps]

  // 构建表头（显示分组标签）
  const headerLabels = ['物料', '物料描述', '基本计量单位']
  for (let i = 1; i <= maxUnits; i++) {
    headerLabels.push(`单位转换${i}`, '', '', '', '')
  }

  let html = `<html><body>`
  html += `<h3>SAP物料单位数据汇总</h3>`
  html += `<p>共 ${data.length} 条记录</p>`
  html += `<table style="border-collapse:collapse;width:100%;font-size:12px;">`

  // 渲染表头
  html += `<tr>`
  headerLabels.forEach((h, idx) => {
    const isSpanning = h === '' ? 'visibility:hidden;' : ''
    html += `<th style="border:1px solid #ddd;padding:6px;background:#f5f5f5;${isSpanning}">${h}</th>`
  })
  html += `</tr>`

  // 渲染数据行（只取前20条）
  const previewData = data.slice(0, 20)
  previewData.forEach(row => {
    html += `<tr>`
    allProps.forEach(prop => {
      html += `<td style="border:1px solid #ddd;padding:6px;">${row[prop] || ''}</td>`
    })
    html += `</tr>`
  })

  html += `</table>`
  html += `<p style="color:#909399;font-size:12px;margin-top:10px;">（此为数据预览，详见附件）</p>`
  html += `</body></html>`

  return html
}
```

---

## 4. 数据格式对比（改造后）

| 数据位置 | 表头格式 | 列内容 | 行数 |
|----------|----------|--------|------|
| 页面表格 | 单层 | 完整列 | 全部 |
| Excel 下载 | 双层合并 | 完整列 | 全部 |
| 邮件附件 Excel | 双层合并 | 完整列 | 全部 |
| 邮件预览 HTML | 分组标签 | 完整列 | 前20行 + 备注 |

---

## 5. 改动文件清单

| 文件 | 改动内容 |
|------|----------|
| `excelProcessor.js` | 1. 新增 `generateAttachmentExcelBuffer`<br>2. 新增 `generateEmailPreviewHtml`<br>3. 新增 `generateEmlContent`<br>4. 删除 `generateAttachmentBase64`（不再需要） |
| `MaterialMergeView.vue` | 1. 导入新增函数<br>2. 改造 `confirmExportEmail` 使用 `generateEmlContent`<br>3. 邮件预览使用 `generateEmailPreviewHtml` |

---

## 6. 依赖

```bash
npm install eml-generator
```

- 零依赖
- 支持浏览器和 Node.js
- 体积小

---

## 7. 测试要点

1. **EML 文件下载** - 下载的 .eml 文件能用 Outlook/Thunderbird 正常打开
2. **附件 Excel 格式** - 打开附件 Excel，表头应为双层合并格式（与"导出Excel"按钮下载的文件一致）
3. **邮件预览 HTML** - 显示完整列（前20行），底部有"此为数据预览，详见附件"备注
4. **中文文件名** - 附件文件名 `material_data.xlsx` 应正确显示（无乱码）
