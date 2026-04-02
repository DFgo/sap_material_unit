import * as XLSX from 'xlsx'

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

  const dataRows = mergedData.map(row =>
    props.map(prop => row[prop] || '')
  )

  const allData = [headerRow1, headerRow2, ...dataRows]
  const ws = XLSX.utils.aoa_to_sheet(allData)

  const merges = []
  let mergeCol = 3

  for (let i = 1; i <= maxUnits; i++) {
    merges.push({
      s: { r: 0, c: mergeCol },
      e: { r: 0, c: mergeCol + 4 }
    })
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
 * 生成邮件附件内容（Base64）- 保留以兼容旧代码
 */
export function generateEmailAttachment(mergedData, maxUnits) {
  const props = ['物料', '物料描述', '基本计量单位']
  for (let i = 1; i <= maxUnits; i++) {
    props.push(`分母${i}`, `基本计量单位${i}`, `可选单位${i}`, `等于${i}`, `计数器${i}`)
  }

  const headerRow = ['物料', '物料描述', '基本计量单位']
  for (let i = 1; i <= maxUnits; i++) {
    headerRow.push(`分母${i}`, `基本计量单位${i}`, `可选单位${i}`, `等于${i}`, `计数器${i}`)
  }

  const dataRows = mergedData.map(row =>
    props.map(prop => row[prop] || '')
  )

  const allData = [headerRow, ...dataRows]
  const ws = XLSX.utils.aoa_to_sheet(allData)
  ws['!cols'] = props.map(() => ({ wch: 12 }))

  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, 'material_data')

  return XLSX.write(wb, { bookType: 'xlsx', type: 'base64' })
}

/**
 * 获取单位转换组的大标题（用于表格显示）
 */
export function getUnitGroupHeaders(maxUnits) {
  const groups = []
  for (let i = 1; i <= maxUnits; i++) {
    groups.push({
      label: `单位转换${i}`,
      cols: [`分母${i}`, `基本计量单位${i}`, `可选单位${i}`, `等于${i}`, `计数器${i}`]
    })
  }
  return groups
}

/**
 * 生成邮件附件 Base64 (不含表头，仅数据)
 */
export function generateAttachmentBase64(data, maxUnits) {
  const props = ['物料', '物料描述', '基本计量单位']
  for (let i = 1; i <= maxUnits; i++) {
    props.push(`分母${i}`, `基本计量单位${i}`, `可选单位${i}`, `等于${i}`, `计数器${i}`)
  }

  const headerRow = ['物料', '物料描述', '基本计量单位']
  for (let i = 1; i <= maxUnits; i++) {
    headerRow.push(`分母${i}`, `基本计量单位${i}`, `可选单位${i}`, `等于${i}`, `计数器${i}`)
  }

  const dataRows = data.map(row => props.map(prop => row[prop] || ''))
  const allData = [headerRow, ...dataRows]
  const ws = XLSX.utils.aoa_to_sheet(allData)
  ws['!cols'] = props.map(() => ({ wch: 12 }))

  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, 'material_data')

  return XLSX.write(wb, { bookType: 'xlsx', type: 'base64' })
}
