import ExcelJS from 'exceljs'

/**
 * 全局表格列定义 - 统一所有表格的列顺序和标签
 * 顺序：物料 → 物料描述 → 基本计量单位 → 单位转换1~N（分母/可选单位/=/计数器/基本计量单位）
 */
const BASE_COLUMNS = [
  { prop: '物料', label: '物料' },
  { prop: '物料描述', label: '物料描述' },
  { prop: '基本计量单位', label: '基本计量单位' },
]

/**
 * 生成指定单位数量的完整列定义（单层表头用）
 * @param {number} maxUnits - 最大单位数
 * @returns {Array<{prop: string, label: string}>} 完整列数组
 */
export function getTableColumns(maxUnits) {
  const columns = [...BASE_COLUMNS]
  for (let i = 1; i <= maxUnits; i++) {
    columns.push(
      { prop: `分母${i}`, label: `分母${i}` },
      { prop: `可选单位${i}`, label: `可选单位${i}` },
      { prop: `等于${i}`, label: `=` },
      { prop: `计数器${i}`, label: `计数器${i}` },
      { prop: `基本计量单位${i}`, label: `基本计量单位${i}` },
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
  return getTableColumns(maxUnits).map((col) => col.prop)
}

/**
 * 生成双层表头的行1（分组大标题）
 * @param {number} maxUnits - 最大单位数
 * @returns {string[]} 表头行1
 */
export function getHeaderRow1(maxUnits) {
  const row1 = ['物料', '物料描述', '基本计量单位']
  for (let i = 1; i <= maxUnits; i++) {
    row1.push(`单位转换${i}`, '', '', '', '')
  }
  return row1
}

/**
 * 生成双层表头的行2（具体列名，与 getTableColumns 顺序一致）
 * @param {number} maxUnits - 最大单位数
 * @returns {string[]} 表头行2
 */
export function getHeaderRow2(maxUnits) {
  const row2 = ['物料', '物料描述', '基本计量单位']
  for (let i = 1; i <= maxUnits; i++) {
    row2.push(`分母${i}`, `可选单位${i}`, `等于${i}`, `计数器${i}`, `基本计量单位${i}`)
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
    reader.onload = async (e) => {
      try {
        const data = new Uint8Array(e.target.result)

        // ExcelJS 异步加载
        const workbook = new ExcelJS.Workbook()
        await workbook.xlsx.load(data)

        const worksheet = workbook.getWorksheet(1)
        const sheetName = worksheet.name

        // 将 worksheet 每行转换为数组（保留空单元格）
        const jsonData = []
        worksheet.eachRow((row) => {
          const values = []
          for (let i = 1; i <= worksheet.columnCount; i++) {
            const cellValue = row.getCell(i).value
            values.push(cellValue === undefined ? '' : cellValue)
          }
          jsonData.push(values)
        })

        const headers = jsonData[0] || []
        const rows = jsonData.slice(1)

        // 过滤掉空行
        const nonEmptyRows = rows.filter((row) => row.some((cell) => cell !== ''))
        const totalRows = nonEmptyRows.length

        // 进度回调（分批处理，避免阻塞）
        let processed = 0
        const batchSize = 50
        const filteredRows = []

        while (processed < totalRows) {
          const end = Math.min(processed + batchSize, totalRows)
          for (let i = processed; i < end; i++) {
            filteredRows.push(nonEmptyRows[i])
            const preview = nonEmptyRows[i].slice(0, 5).join(' | ')
            onProgress(i + 1, totalRows, preview)
          }
          processed = end
          // 让UI有机会更新
          await new Promise((r) => setTimeout(r, 0))
        }

        onProgress(totalRows, totalRows, '处理完成')
        resolve({ headers, data: filteredRows, sheetName })
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
    reader.onload = async (e) => {
      try {
        const data = new Uint8Array(e.target.result)

        const workbook = new ExcelJS.Workbook()
        await workbook.xlsx.load(data)

        const worksheet = workbook.getWorksheet(1)
        const sheetName = worksheet.name

        const jsonData = []
        worksheet.eachRow((row) => {
          const values = []
          for (let i = 1; i <= worksheet.columnCount; i++) {
            const cellValue = row.getCell(i).value
            values.push(cellValue === undefined ? '' : cellValue)
          }
          jsonData.push(values)
        })

        const headers = jsonData[0] || []
        const rows = jsonData.slice(1).filter((row) => row.some((cell) => cell !== ''))

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
  return rows
    .map((row) => ({
      物料: row[0] || '',
      物料描述: row[row.length - 1] || '',
      基本计量单位: row[12] || '',
    }))
    .filter((row) => row.物料 !== '')
}

/**
 * 解析 material_unit 文件
 * 取列: 第0、1、2、3列
 */
export function parseMaterialUnit(rawData) {
  const rows = rawData.data
  return rows
    .map((row) => ({
      物料: row[0] || '',
      可选计量单位: row[1] || '',
      计数器: row[2] || '',
      分母: row[3] || '',
    }))
    .filter((row) => row.物料 !== '')
}

/**
 * 合并数据，添加单位转换列
 */
export function processMerge(materialData, materialUnit) {
  const unitCounts = {}
  materialUnit.forEach((row) => {
    if (!unitCounts[row.物料]) {
      unitCounts[row.物料] = []
    }
    unitCounts[row.物料].push(row)
  })

  const maxUnits = Math.max(...Object.values(unitCounts).map((v) => v.length), 0)

  const headers = [
    { label: '物料', prop: '物料', span: 1 },
    { label: '物料描述', prop: '物料描述', span: 1 },
    { label: '基本计量单位', prop: '基本计量单位', span: 1 },
  ]

  for (let i = 1; i <= maxUnits; i++) {
    headers.push(
      { label: `分母${i}`, prop: `分母${i}`, span: 1 },
      { label: `基本计量单位${i}`, prop: `基本计量单位${i}`, span: 1 },
      { label: `等于${i}`, prop: `等于${i}`, span: 1 },
      { label: `计数器${i}`, prop: `计数器${i}`, span: 1 },
      { label: `可选单位${i}`, prop: `可选单位${i}`, span: 1 },
    )
  }

  const mergedData = materialData.map((row) => {
    const newRow = {
      物料: row.物料,
      物料描述: row.物料描述,
      基本计量单位: row.基本计量单位,
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
 * 通用构建 Excel 工作簿（导出 + 邮件附件 共用）
 * 复用所有样式、合并、格式
 */
async function buildExcelWorkbook(mergedData, maxUnits) {
  const props = getColumnProps(maxUnits)
  const headerRow1 = getHeaderRow1(maxUnits)
  const headerRow2 = getHeaderRow2(maxUnits)
  const dataRows = mergedData.map((row) => props.map((prop) => row[prop] || ''))
  const allData = [headerRow1, headerRow2, ...dataRows]

  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet('物料单位转换数据')

  // 插入所有行
  allData.forEach(rowData => {
    worksheet.addRow(rowData)
  })

  // 前3列垂直合并
  worksheet.mergeCells('A1:A2')
  worksheet.mergeCells('B1:B2')
  worksheet.mergeCells('C1:C2')

  // 单位转换组合并
  let startCol = 4
  for (let i = 1; i <= maxUnits; i++) {
    worksheet.mergeCells(1, startCol, 1, startCol + 4)
    const cell = worksheet.getCell(1, startCol)
    cell.alignment = { horizontal: 'center', vertical: 'center', wrapText: true }
    startCol += 5
  }

  // 全局样式：居中 + 边框
  worksheet.eachRow({ includeEmpty: false }, (row) => {
    row.eachCell({ includeEmpty: true }, (cell) => {
      cell.alignment = { horizontal: 'center', vertical: 'center', wrapText: true }
      cell.border = {
        top: { style: 'thin' },
        left: { style: 'thin' },
        bottom: { style: 'thin' },
        right: { style: 'thin' },
      }
    })
  })

  // 列宽 + 行高
  worksheet.getColumn('B').width = 70
  worksheet.getRow(1).height = 30

  return workbook
}

/**
 * 导出为 Excel 文件（带合并标题行）
 */
export async function exportToExcel(mergedData, maxUnits, filename = 'merged_output.xlsx') {
  // const props = getColumnProps(maxUnits)
  // const headerRow1 = getHeaderRow1(maxUnits)
  // const headerRow2 = getHeaderRow2(maxUnits)
  // const dataRows = mergedData.map((row) => props.map((prop) => row[prop] || ''))
  // const allData = [headerRow1, headerRow2, ...dataRows]

  // // 创建工作簿和工作表
  // const workbook = new ExcelJS.Workbook()
  // const worksheet = workbook.addWorksheet('物料单位转换数据')

  // // 先加所有行
  // allData.forEach((rowData) => {
  //   worksheet.addRow(rowData)
  // })

  // // 合并前3列的两行表头（物料/物料描述/基本计量单位）
  // worksheet.mergeCells('A1:A2')
  // worksheet.mergeCells('B1:B2')
  // worksheet.mergeCells('C1:C2')

  // // 单位转换组合并
  // let startCol = 4 // D列开始
  // for (let i = 1; i <= maxUnits; i++) {
  //   // 直接拼接单元格：第1行，startCol ~ startCol+4
  //   worksheet.mergeCells(1, startCol, 1, startCol + 4)

  //   const cell = worksheet.getCell(1, startCol)
  //   cell.alignment = {
  //     horizontal: 'center',
  //     vertical: 'center',
  //     wrapText: true,
  //   }

  //   startCol += 5
  // }

  // // 全局样式：居中 + 全边框
  // worksheet.eachRow({ includeEmpty: false }, (row) => {
  //   row.eachCell({ includeEmpty: true }, (cell) => {
  //     cell.alignment = { horizontal: 'center', vertical: 'center', wrapText: true }
  //     cell.border = {
  //       top: { style: 'thin' },
  //       left: { style: 'thin' },
  //       bottom: { style: 'thin' },
  //       right: { style: 'thin' },
  //     }
  //   })
  // })

  // // 物料描述-B列设置宽度
  // worksheet.getColumn('B').width = 70

  // // 第一列设置高度
  // worksheet.getRow(1).height = 30

  // // 调用下载工具函数
  // await downloadExcelFile(workbook, filename)
  const workbook = await buildExcelWorkbook(mergedData, maxUnits)
  await downloadExcelFile(workbook, filename)
}

function getFileName(filename) {
  const timestamp = new Date()
    .toISOString()
    .replace(/[-:\.T]/g, '')
    .substring(0, 14)
  const randomStr = Math.random().toString(36).substring(2, 6).toUpperCase()
  const nameParts = filename.split('.')
  const ext = nameParts.pop()
  const finalName = `${nameParts.join('.')}_${timestamp}_${randomStr}.${ext}`
  return finalName
}

async function downloadExcelFile(workbook, filename) {
  const buffer = await workbook.xlsx.writeBuffer()
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  })
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = getFileName(filename)
  link.click()
  URL.revokeObjectURL(url)
}

/**
 * 生成预览表格 HTML（复用 getTableColumns 统一列顺序）
 */
export function generatePreviewTable(data, maxUnits) {
  if (!data || data.length === 0) return ''

  const columns = getTableColumns(maxUnits)

  let html = '<table style="border-collapse:collapse;width:100%;font-size:12px;">'
  html += '<tr>'
  columns.forEach((col) => {
    html += `<th style="border:1px solid #ddd;padding:6px;background:#f5f5f5;">${col.label}</th>`
  })
  html += '</tr>'

  data.slice(0, 20).forEach((row) => {
    html += '<tr>'
    columns.forEach((col) => {
      html += `<td style="border:1px solid #ddd;padding:6px;">${row[col.prop] || ''}</td>`
    })
    html += '</tr>'
  })

  html += '</table>'
  return html
}

/**
 * 生成邮件附件 Excel Base64 字符串（双层表头，与 exportToExcel 格式一致）
 * @param {Array} mergedData - 合并后数据
 * @param {number} maxUnits - 最大单位数
 * @returns {Promise<string>} Excel 文件 Base64 字符串
 */
export async function generateAttachmentExcelBuffer(mergedData, maxUnits) {
  // const props = getColumnProps(maxUnits)
  // const headerRow1 = getHeaderRow1(maxUnits)
  // const headerRow2 = getHeaderRow2(maxUnits)
  // const dataRows = mergedData.map((row) => props.map((prop) => row[prop] || ''))
  // const allData = [headerRow1, headerRow2, ...dataRows]

  // const workbook = new ExcelJS.Workbook()
  // const worksheet = workbook.addWorksheet('material_data')

  // // 先加所有行，不再循环里做合并
  // allData.forEach((rowData) => {
  //   worksheet.addRow(rowData)
  // })

  // // 统一合并（移出循环）
  // let mergeCol = 3
  // for (let i = 1; i <= maxUnits; i++) {
  //   worksheet.mergeCells(1, mergeCol + 1, 1, mergeCol + 5)
  //   const cell = worksheet.getCell(1, mergeCol + 1)
  //   cell.alignment = { horizontal: 'center', vertical: 'center', wrapText: true }
  //   mergeCol += 5
  // }

  // props.forEach((_, idx) => {
  //   worksheet.getColumn(idx + 1).width = idx < 3 ? 15 : 12
  // })

  // const excelBuffer = await workbook.xlsx.writeBuffer()

  // // 换成正确的 base64
  // return bufferToBase64(excelBuffer)
  const workbook = await buildExcelWorkbook(mergedData, maxUnits)
  const buffer = await workbook.xlsx.writeBuffer()
  return bufferToBase64(buffer)
}


function bufferToBase64(buffer) {
  return new Promise((resolve) => {
    const blob = new Blob([buffer], { type: 'application/octet-stream' })
    const reader = new FileReader()
    reader.onloadend = () => resolve(reader.result.split(',')[1])
    reader.readAsDataURL(blob)
  })
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
  const headerRow2 = getHeaderRow2(maxUnits)

  let html = `<html><body>`
  html += `<h3>SAP物料单位数据汇总</h3>`
  html += `<p>共 ${data.length} 条记录</p>`
  html += `<table style="border-collapse:collapse;width:100%;font-size:12px;">`

  // 渲染表头第1行（分组大标题）
  html += `<tr>`
  headerRow1.forEach((h) => {
    const style = h === '' ? 'visibility:hidden;' : ''
    html += `<th style="border:1px solid #ddd;padding:6px;background:#f5f5f5;${style}">${h}</th>`
  })
  html += `</tr>`

  // 渲染表头第2行（具体列名）
  html += `<tr>`
  headerRow2.forEach((h) => {
    const style = h === '' ? 'visibility:hidden;' : ''
    html += `<th style="border:1px solid #ddd;padding:6px;background:#f5f5f5;${style}">${h}</th>`
  })
  html += `</tr>`

  // 渲染数据行（只取前20条）
  const previewData = data.slice(0, 20)
  previewData.forEach((row) => {
    html += `<tr>`
    columns.forEach((col) => {
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
 * 生成 EML 格式字符串（自定义 MIME 拼接）
 * @param {Array} mergedData - 合并后数据
 * @param {number} maxUnits - 最大单位数
 * @param {Object} emailOptions - 邮件选项 { to, from, subject }
 * @returns {Promise<string>} EML 格式字符串
 */
export async function generateEmlContent(mergedData, maxUnits, emailOptions) {
  const config = {
    from: 'sap-system@company.com',
    subject: 'SAP物料单位数据',
    ...emailOptions,
  }

  const [excelBase64, previewHtml] = await Promise.all([
    generateAttachmentExcelBuffer(mergedData, maxUnits),
    generateEmailPreviewHtml(mergedData, maxUnits),
  ])

  const boundary = `----=_Part_${Math.random().toString(36).slice(2)}`
  const date = new Date().toUTCString()
  const base64Formatted = excelBase64.replace(/.{76}/g, '$&\n')

  const emlLines = [
    'MIME-Version: 1.0',
    `From: ${config.from}`,
    `To: ${config.to}`,
    `Subject: ${config.subject}`,
    `Date: ${date}`,
    `Content-Type: multipart/mixed; boundary="${boundary}"`,
    '',
    `--${boundary}`,
    'Content-Type: text/html; charset=utf-8',
    '',
    previewHtml,
    '',
    `--${boundary}`,
    'Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet; name="material_data.xlsx"',
    'Content-Transfer-Encoding: base64',
    'Content-Disposition: attachment; filename="material_data.xlsx"',
    '',
    base64Formatted,
    '',
    `--${boundary}--`,
  ]

  return emlLines.join('\n')
}
