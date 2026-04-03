import * as XLSX from 'xlsx'
import { eml } from 'eml-generator'

/**
 * 全局表格列定义 - 统一所有表格的列顺序和标签
 * 顺序：物料 → 物料描述 → 基本计量单位 → 单位转换1~N（分母/基本计量单位/可选单位/=/计数器）
 */
const BASE_COLUMNS = [
  { prop: '物料', label: '物料' },
  { prop: '物料描述', label: '物料描述' },
  { prop: '基本计量单位', label: '基本计量单位' }
]

/**
 * 生成指定单位数量的完整列定义
 * @param {number} maxUnits - 最大单位数
 * @returns {Array<{prop: string, label: string}>} 完整列数组
 */
export function getTableColumns(maxUnits) {
  const columns = [...BASE_COLUMNS]
  for (let i = 1; i <= maxUnits; i++) {
    columns.push(
      { prop: `分母${i}`, label: `分母${i}` },
      { prop: `基本计量单位${i}`, label: `基本计量单位${i}` },
      { prop: `等于${i}`, label: `=` },
      { prop: `计数器${i}`, label: `计数器${i}` },
      { prop: `可选单位${i}`, label: `可选单位${i}` },
    )
  }
  return columns
}

/**
 * 获取列的 prop 数组（按全局定义顺序）
 * @param {number} maxUnits - 最大单位数
 * @returns {string[]} prop 数组
 */
export function getColumnProps(maxUnits) {
  return getTableColumns(maxUnits).map(col => col.prop)
}

/**
 * 生成双层表头的行1（分组大标题）
 * @param {number} maxUnits - 最大单位数
 * @returns {string[]} 表头行1
 */
function getHeaderRow1(maxUnits) {
  const row1 = ['物料', '物料描述', '基本计量单位']
  for (let i = 1; i <= maxUnits; i++) {
    row1.push(`单位转换${i}`, '', '', '', '')
  }
  return row1
}

/**
 * 生成双层表头的行2（具体列名）
 * @param {number} maxUnits - 最大单位数
 * @returns {string[]} 表头行2
 */
function getHeaderRow2(maxUnits) {
  const row2 = ['物料', '物料描述', '基本计量单位']
  for (let i = 1; i <= maxUnits; i++) {
    row2.push(`分母${i}`, `基本计量单位${i}`, `可选单位${i}`, `等于${i}`, `计数器${i}`)
  }
  return row2
}

/**
 * 读取 Excel 文件（带进度回调）
 * @param {File} file - Excel 文件
 * @param {Function} onProgress - 进度回调 (currentRow, totalRows, rowContent)
 * @returns {Promise<{headers: string[], data: any[][], sheetName: string}>}
 */
export async function readExcelFileWithProgress(file, onProgress) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result)
        const workbook = XLSX.read(data, { type: 'array', cellDates: true })
        const sheetName = workbook.SheetNames[0]
        const sheet = workbook.Sheets[sheetName]

        // 转换为数组格式（保留所有单元格，包括空单元格）
        const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' })

        const headers = jsonData[0] || []
        const rows = jsonData.slice(1)

        // 过滤掉空行
        const nonEmptyRows = rows.filter(row => row.some(cell => cell !== ''))
        const totalRows = nonEmptyRows.length

        // 进度回调（分批处理，避免阻塞）
        let processed = 0
        const batchSize = 50
        const filteredRows = []

        function processBatch() {
          return new Promise((resolve) => {
            const end = Math.min(processed + batchSize, totalRows)
            for (let i = processed; i < end; i++) {
              filteredRows.push(nonEmptyRows[i])
              const preview = nonEmptyRows[i].slice(0, 5).join(' | ')
              onProgress(i + 1, totalRows, preview)
            }
            processed = end
            resolve()
          })
        }

        async function processAll() {
          while (processed < totalRows) {
            await processBatch()
            // 让UI有机会更新
            await new Promise(r => setTimeout(r, 0))
          }

          onProgress(totalRows, totalRows, '处理完成')
          resolve({ headers, data: filteredRows, sheetName })
        }

        processAll().catch(reject)
      } catch (err) {
        reject(err)
      }
    }
    reader.onerror = reject
    reader.readAsArrayBuffer(file)
  })
}

/**
 * 读取 Excel 文件（无进度，纯读取）
 */
export async function readExcelFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result)
        const workbook = XLSX.read(data, { type: 'array', cellDates: true })
        const sheetName = workbook.SheetNames[0]
        const sheet = workbook.Sheets[sheetName]

        const jsonData = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' })

        const headers = jsonData[0] || []
        const rows = jsonData.slice(1).filter(row => row.some(cell => cell !== ''))

        resolve({ headers, data: rows, sheetName })
      } catch (err) {
        reject(err)
      }
    }
    reader.onerror = reject
    reader.readAsArrayBuffer(file)
  })
}

/**
 * 解析 material_data 文件
 * 取列: 第0列(物料), 倒数第1列, 第12列
 */
export function parseMaterialData(rawData) {
  const rows = rawData.data
  return rows.map(row => ({
    物料: row[0] || '',
    物料描述: row[row.length - 1] || '',
    基本计量单位: row[12] || ''
  })).filter(row => row.物料 !== '')
}

/**
 * 解析 material_unit 文件
 * 取列: 第0、1、2、3列
 */
export function parseMaterialUnit(rawData) {
  const rows = rawData.data
  return rows.map(row => ({
    物料: row[0] || '',
    可选计量单位: row[1] || '',
    计数器: row[2] || '',
    分母: row[3] || ''
  })).filter(row => row.物料 !== '')
}

/**
 * 合并数据，添加单位转换列
 */
export function processMerge(materialData, materialUnit) {
  const unitCounts = {}
  materialUnit.forEach(row => {
    if (!unitCounts[row.物料]) {
      unitCounts[row.物料] = []
    }
    unitCounts[row.物料].push(row)
  })

  const maxUnits = Math.max(...Object.values(unitCounts).map(v => v.length), 0)

  const headers = [
    { label: '物料', prop: '物料', span: 1 },
    { label: '物料描述', prop: '物料描述', span: 1 },
    { label: '基本计量单位', prop: '基本计量单位', span: 1 }
  ]

  for (let i = 1; i <= maxUnits; i++) {
    headers.push(
      { label: `分母${i}`, prop: `分母${i}`, span: 1 },
      { label: `基本计量单位${i}`, prop: `基本计量单位${i}`, span: 1 },
      { label: `可选单位${i}`, prop: `可选单位${i}`, span: 1 },
      { label: `等于${i}`, prop: `等于${i}`, span: 1 },
      { label: `计数器${i}`, prop: `计数器${i}`, span: 1 }
    )
  }

  const mergedData = materialData.map(row => {
    const newRow = {
      物料: row.物料,
      物料描述: row.物料描述,
      基本计量单位: row.基本计量单位
    }

    const units = unitCounts[row.物料] || []
    units.forEach((u, idx) => {
      const i = idx + 1
      newRow[`分母${i}`] = u.分母
      newRow[`基本计量单位${i}`] = row.基本计量单位
      newRow[`可选单位${i}`] = u.可选计量单位
      newRow[`等于${i}`] = '='
      newRow[`计数器${i}`] = u.计数器
    })

    for (let i = units.length + 1; i <= maxUnits; i++) {
      newRow[`分母${i}`] = ''
      newRow[`基本计量单位${i}`] = ''
      newRow[`可选单位${i}`] = ''
      newRow[`等于${i}`] = ''
      newRow[`计数器${i}`] = ''
    }

    return newRow
  })

  return { mergedData, maxUnits, headers }
}

/**
 * 导出为 Excel 文件（带合并标题行）
 */
export function exportToExcel(mergedData, maxUnits, filename = 'merged_output.xlsx') {
  const props = getColumnProps(maxUnits)
  const headerRow1 = getHeaderRow1(maxUnits)
  const headerRow2 = getHeaderRow2(maxUnits)

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
  XLSX.writeFile(wb, filename)
}

/**
 * 生成预览表格 HTML
 */
export function generatePreviewTable(data, maxUnits) {
  if (!data || data.length === 0) return ''

  const props = ['物料', '物料描述', '基本计量单位']
  for (let i = 1; i <= maxUnits; i++) {
    props.push(`分母${i}`, `基本计量单位${i}`, `可选单位${i}`, `等于${i}`, `计数器${i}`)
  }

  const headers = ['物料', '物料描述', '基本计量单位']

  let html = '<table style="border-collapse:collapse;width:100%;font-size:12px;">'
  html += '<tr>'
  headers.forEach(h => {
    html += `<th style="border:1px solid #ddd;padding:6px;background:#f5f5f5;">${h}</th>`
  })
  html += '</tr>'

  data.slice(0, 20).forEach(row => {
    html += '<tr>'
    headers.forEach(prop => {
      html += `<td style="border:1px solid #ddd;padding:6px;">${row[prop] || ''}</td>`
    })
    html += '</tr>'
  })

  html += '</table>'
  return html
}

/**
 * 生成邮件附件 Excel Buffer（双层表头，与 exportToExcel 格式一致）
 * @param {Array} mergedData - 合并后数据
 * @param {number} maxUnits - 最大单位数
 * @returns {Buffer} Excel 文件 Buffer
 */
export function generateAttachmentExcelBuffer(mergedData, maxUnits) {
  const props = getColumnProps(maxUnits)
  const headerRow1 = getHeaderRow1(maxUnits)
  const headerRow2 = getHeaderRow2(maxUnits)

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

/**
 * 生成邮件预览 HTML（完整列，前20行，带预览备注）
 * @param {Array} data - 数据（通常传前20条）
 * @param {number} maxUnits - 最大单位数
 * @returns {string} HTML 字符串
 */
export function generateEmailPreviewHtml(data, maxUnits) {
  if (!data || data.length === 0) return ''

  const columns = getTableColumns(maxUnits)
  const headerRow1 = getHeaderRow1(maxUnits)

  let html = `<html><body>`
  html += `<h3>SAP物料单位数据汇总</h3>`
  html += `<p>共 ${data.length} 条记录</p>`
  html += `<table style="border-collapse:collapse;width:100%;font-size:12px;">`

  // 渲染表头
  html += `<tr>`
  headerRow1.forEach((h) => {
    const style = h === '' ? 'visibility:hidden;' : ''
    html += `<th style="border:1px solid #ddd;padding:6px;background:#f5f5f5;${style}">${h}</th>`
  })
  html += `</tr>`

  // 渲染数据行（只取前20条）
  const previewData = data.slice(0, 20)
  previewData.forEach(row => {
    html += `<tr>`
    columns.forEach(col => {
      html += `<td style="border:1px solid #ddd;padding:6px;">${row[col.prop] || ''}</td>`
    })
    html += `</tr>`
  })

  html += `</table>`
  html += `<p style="color:#909399;font-size:12px;margin-top:10px;">（此为数据预览，详见附件）</p>`
  html += `</body></html>`

  return html
}

/**
 * 生成 EML 格式字符串（使用 eml-generator）
 * @param {Array} mergedData - 合并后数据
 * @param {number} maxUnits - 最大单位数
 * @param {Object} emailOptions - 邮件选项 { to, from, subject }
 * @returns {string} EML 格式字符串
 */
export function generateEmlContent(mergedData, maxUnits, emailOptions) {
  // 生成附件 Excel Buffer（双层表头）
  const excelBuffer = generateAttachmentExcelBuffer(mergedData, maxUnits)

  // 将 Buffer 转为 Base64 字符串（分块处理避免栈溢出）
  const uint8Array = new Uint8Array(excelBuffer)
  const chunkSize = 8192
  let base64Excel = ''
  for (let i = 0; i < uint8Array.length; i += chunkSize) {
    const chunk = uint8Array.slice(i, i + chunkSize)
    base64Excel += String.fromCharCode.apply(null, chunk)
  }
  base64Excel = btoa(base64Excel)

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
        data: base64Excel,
        contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      }
    ]
  })

  return emailContent
}
