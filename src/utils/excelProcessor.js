import * as XLSX from 'xlsx'

/**
 * 读取 Excel 文件
 * @param {File} file - Excel 文件
 * @returns {Promise<{headers: string[], data: any[][], sheetName: string}>}
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

        // 转换为数组格式（保留所有单元格，包括空单元格）
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
  // 取第0列、倒数第1列（-1即length-1）、第12列
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
 * 返回 { mergedData: [], maxUnits: number, headers: [] }
 */
export function processMerge(materialData, materialUnit) {
  // 获取最大单位转换数量
  const unitCounts = {}
  materialUnit.forEach(row => {
    if (!unitCounts[row.物料]) {
      unitCounts[row.物料] = []
    }
    unitCounts[row.物料].push(row)
  })

  const maxUnits = Math.max(...Object.values(unitCounts).map(v => v.length), 0)
  console.log(`检测到单个物料最多有 ${maxUnits} 个单位转换`)

  // 构建表头
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

  // 构建合并后的数据
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

    // 填充空列
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
  // 构建工作表数据
  // 第1行: 大标题（物料、物料描述、基本计量单位 + 单位转换1-5合并标题）
  // 第2行: 小标题

  const props = ['物料', '物料描述', '基本计量单位']
  for (let i = 1; i <= maxUnits; i++) {
    props.push(`分母${i}`, `基本计量单位${i}`, `可选单位${i}`, `等于${i}`, `计数器${i}`)
  }

  // 构建表头行
  const headerRow1 = ['物料', '物料描述', '基本计量单位']
  const headerRow2 = ['物料', '物料描述', '基本计量单位']
  let col = 4

  for (let i = 1; i <= maxUnits; i++) {
    headerRow1.push(`单位转换${i}`, '', '', '', '') // 合并5列
    headerRow2.push(`分母${i}`, `基本计量单位${i}`, `可选单位${i}`, `等于${i}`, `计数器${i}`)
    col += 5
  }

  // 构建数据行
  const dataRows = mergedData.map(row =>
    props.map(prop => row[prop] || '')
  )

  // 合并所有数据
  const allData = [headerRow1, headerRow2, ...dataRows]

  // 创建工作表
  const ws = XLSX.utils.aoa_to_sheet(allData)

  // 设置合并单元格
  const merges = []
  let mergeCol = 3 // 从第4列开始（索引3）

  for (let i = 1; i <= maxUnits; i++) {
    merges.push({
      s: { r: 0, c: mergeCol },
      e: { r: 0, c: mergeCol + 4 }
    })
    mergeCol += 5
  }

  ws['!merges'] = merges

  // 设置列宽
  ws['!cols'] = props.map((_, idx) => {
    if (idx < 3) return { wch: 15 }
    return { wch: 12 }
  })

  // 创建工作簿
  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, 'material_data')

  // 写入文件
  XLSX.writeFile(wb, filename)
}

/**
 * 生成邮件附件内容（Base64）
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

  // 合并单元格（只有数据区域的合并，没有大标题）
  ws['!cols'] = props.map((_, idx) => {
    if (idx < 3) return { wch: 15 }
    return { wch: 12 }
  })

  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, 'material_data')

  // 生成 base64
  const wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'base64' })
  return wbout
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
