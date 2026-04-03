<script setup>
import { ref, computed } from 'vue'
import { ElUpload, ElButton, ElTable, ElTableColumn, ElInput, ElMessage, ElCard, ElDialog, ElTabs, ElTabPane, ElPagination } from 'element-plus'
import {
  readExcelFileWithProgress,
  parseMaterialData,
  parseMaterialUnit,
  processMerge,
  exportToExcel,
  generateEmlContent,
  generateEmailPreviewHtml,
  getTableColumns
} from '../utils/excelProcessor'

// 状态
const activeTab = ref('upload')
const materialDataFile = ref(null)
const materialUnitFile = ref(null)
const materialData = ref([])
const materialUnit = ref([])
const mergedData = ref([])
const rawMaterialData = ref([])  // 原始数据（未解析）
const rawMaterialUnit = ref([])  // 原始数据（未解析）
const maxUnits = ref(0)
const tableHeaders = ref([])
const searchQuery = ref('')

// 原始数据分页
const rawDataPage = ref(1)
const rawDataPageSize = ref(100)

// 上传进度
const uploadProgressVisible = ref(false)
const uploadProgress = ref({
  fileName: '',
  currentRow: 0,
  totalRows: 0,
  currentContent: '',
  percent: 0
})

// 分页
const currentPage = ref(1)
const pageSize = ref(50)

// 邮件对话框
const emailDialogVisible = ref(false)
const emailForm = ref({
  to: '',
  subject: 'SAP物料单位数据'
})
const emailPreviewHtml = computed(() => generateEmailPreviewHtml(mergedData.value.slice(0, 20), maxUnits.value))

// 过滤后的数据
const filteredData = computed(() => {
  if (!searchQuery.value) return mergedData.value
  const q = searchQuery.value.toLowerCase()
  return mergedData.value.filter(row =>
    row.物料?.toString().toLowerCase().includes(q) ||
    row.物料描述?.toString().toLowerCase().includes(q)
  )
})

// 分页后的数据
const paginatedData = computed(() => {
  const start = (currentPage.value - 1) * pageSize.value
  const end = start + pageSize.value
  return filteredData.value.slice(start, end)
})

// 表格列（minWidth 仅用于页面展示，全局列顺序由 getTableColumns 统一控制）
const tableColumns = computed(() => {
  const minWidthMap = {
    '物料': 100,
    '物料描述': 150,
    '基本计量单位': 90,
    '=': 50
  }
  return getTableColumns(maxUnits.value).map(col => ({
    ...col,
    minWidth: minWidthMap[col.label] || 80
  }))
})

// 源数据列（原始数据 - 根据第一行数据动态计算）
const rawMaterialDataColumns = computed(() => {
  if (!rawMaterialData.value.length) return []
  const firstRow = rawMaterialData.value[0]
  const maxCols = Object.keys(firstRow).filter(k => k.startsWith('col')).length
  const displayCols = Math.min(maxCols, 15)
  return Array.from({ length: displayCols }, (_, i) => ({
    prop: `col${i}`,
    label: `列${i + 1}`,
    minWidth: 80
  }))
})

const rawMaterialUnitColumns = computed(() => {
  if (!rawMaterialUnit.value.length) return []
  const firstRow = rawMaterialUnit.value[0]
  const maxCols = Object.keys(firstRow).filter(k => k.startsWith('col')).length
  const displayCols = Math.min(maxCols, 15)
  return Array.from({ length: displayCols }, (_, i) => ({
    prop: `col${i}`,
    label: `列${i + 1}`,
    minWidth: 80
  }))
})

// 源数据分页后的数据
const paginatedRawMaterialData = computed(() => {
  const start = (rawDataPage.value - 1) * rawDataPageSize.value
  const end = start + rawDataPageSize.value
  return rawMaterialData.value.slice(start, end)
})

const rawMaterialDataTotal = computed(() => rawMaterialData.value.length)

function handleRawDataPageChange(page) {
  rawDataPage.value = page
}

// 文件上传处理（带进度）
async function handleMaterialDataUpload(file) {
  materialDataFile.value = file.raw
  uploadProgress.value = {
    fileName: file.name,
    currentRow: 0,
    totalRows: 0,
    currentContent: '',
    percent: 0
  }
  uploadProgressVisible.value = true

  try {
    const result = await readExcelFileWithProgress(file.raw, (row, total, content) => {
      uploadProgress.value.currentRow = row
      uploadProgress.value.totalRows = total
      uploadProgress.value.currentContent = content
      uploadProgress.value.percent = Math.round((row / total) * 100)
    })

    rawMaterialData.value = result.data.map((row, idx) => {
      const obj = { _index: idx }
      result.headers.forEach((_, i) => {
        obj[`col${i}`] = row[i]
      })
      return obj
    })

    materialData.value = parseMaterialData(result)
    uploadProgressVisible.value = false
    ElMessage.success(`material_data.xlsx 读取成功，共 ${materialData.value.length} 条数据`)
    checkReady()
  } catch (err) {
    uploadProgressVisible.value = false
    ElMessage.error('读取 material_data.xlsx 失败')
    console.error(err)
  }
  return false
}

async function handleMaterialUnitUpload(file) {
  materialUnitFile.value = file.raw
  uploadProgress.value = {
    fileName: file.name,
    currentRow: 0,
    totalRows: 0,
    currentContent: '',
    percent: 0
  }
  uploadProgressVisible.value = true

  try {
    const result = await readExcelFileWithProgress(file.raw, (row, total, content) => {
      uploadProgress.value.currentRow = row
      uploadProgress.value.totalRows = total
      uploadProgress.value.currentContent = content
      uploadProgress.value.percent = Math.round((row / total) * 100)
    })

    rawMaterialUnit.value = result.data.map((row, idx) => {
      const obj = { _index: idx }
      result.headers.forEach((_, i) => {
        obj[`col${i}`] = row[i]
      })
      return obj
    })

    materialUnit.value = parseMaterialUnit(result)
    uploadProgressVisible.value = false
    ElMessage.success(`material_unit.xlsx 读取成功，共 ${materialUnit.value.length} 条数据`)
    checkReady()
  } catch (err) {
    uploadProgressVisible.value = false
    ElMessage.error('读取 material_unit.xlsx 失败')
    console.error(err)
  }
  return false
}

// 检查是否准备好计算
const canCalculate = computed(() => {
  return materialData.value.length > 0 && materialUnit.value.length > 0
})

function checkReady() {
  if (canCalculate.value) {
    ElMessage.info('两个文件已上传完成，点击"开始计算"进行数据合并')
  }
}

// 开始计算
function startCalculate() {
  if (!canCalculate.value) {
    ElMessage.warning('请先上传两个文件')
    return
  }

  try {
    const result = processMerge(materialData.value, materialUnit.value)
    mergedData.value = result.mergedData
    maxUnits.value = result.maxUnits
    tableHeaders.value = result.headers
    currentPage.value = 1
    activeTab.value = 'merged'
    ElMessage.success(`合并完成！共 ${mergedData.value.length} 条数据，最大 ${maxUnits.value} 个单位转换`)
  } catch (err) {
    ElMessage.error('合并失败')
    console.error(err)
  }
}

// 导出 Excel
async function handleExportExcel() {
  if (mergedData.value.length === 0) {
    ElMessage.warning('没有数据可导出')
    return
  }
  try {
    await exportToExcel(mergedData.value, maxUnits.value, 'merged_output.xlsx')
    ElMessage.success('Excel 导出成功')
  } catch (err) {
    ElMessage.error('导出失败')
    console.error(err)
  }
}

// 导出 EML 邮件
function handleExportEmail() {
  if (mergedData.value.length === 0) {
    ElMessage.warning('没有数据可导出')
    return
  }
  emailDialogVisible.value = true
}

async function confirmExportEmail() {
  if (!emailForm.value.to) {
    ElMessage.warning('请输入收件人邮箱')
    return
  }

  try {
    const emlContent = await generateEmlContent(
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

// 删除行
function deleteRow(index) {
  const realIndex = (currentPage.value - 1) * pageSize.value + index
  mergedData.value.splice(realIndex, 1)
  ElMessage.success('已删除')
}

// 添加行
function addRow() {
  const newRow = { 物料: '', 物料描述: '', 基本计量单位: '' }
  for (let i = 1; i <= maxUnits.value; i++) {
    newRow[`分母${i}`] = ''
    newRow[`基本计量单位${i}`] = ''
    newRow[`可选单位${i}`] = ''
    newRow[`等于${i}`] = ''
    newRow[`计数器${i}`] = ''
  }
  mergedData.value.push(newRow)
}

// 页码改变
function handlePageChange(page) {
  currentPage.value = page
}
</script>

<template>
  <div class="container">
    <el-tabs v-model="activeTab" class="main-tabs">
      <!-- 上传页 -->
      <el-tab-pane label="上传文件" name="upload">
        <el-card class="upload-card">
          <template #header>
            <div class="card-header">
              <span>上传 Excel 文件</span>
            </div>
          </template>

          <div class="upload-section">
            <div class="upload-item">
              <h4>material_data.xlsx <span class="source-tag">来源: SAP SE16N MARA表</span></h4>
              <p class="upload-hint">取第1列、倒数第1列、第13列</p>
              <el-upload
                :auto-upload="false"
                :show-file-list="true"
                :on-change="handleMaterialDataUpload"
                accept=".xlsx,.xls"
              >
                <el-button type="primary">选择文件</el-button>
              </el-upload>
            </div>

            <div class="upload-item">
              <h4>material_unit.xlsx <span class="source-tag">来源: SAP SE16N MARM表</span></h4>
              <p class="upload-hint">取第1-4列（物料、可选计量单位、计数器、分母）</p>
              <el-upload
                :auto-upload="false"
                :show-file-list="true"
                :on-change="handleMaterialUnitUpload"
                accept=".xlsx,.xls"
              >
                <el-button type="primary">选择文件</el-button>
              </el-upload>
            </div>
          </div>

          <div class="upload-status">
            <el-tag v-if="materialData.length > 0" type="success" size="large">material_data: {{ materialData.length }} 条</el-tag>
            <el-tag v-if="materialUnit.length > 0" type="success" size="large">material_unit: {{ materialUnit.length }} 条</el-tag>
            <el-tag v-if="canCalculate" type="warning" size="large">待计算</el-tag>
          </div>

          <div class="calculate-btn-wrapper">
            <el-button type="warning" size="large" :disabled="!canCalculate" @click="startCalculate">
              开始计算
            </el-button>
          </div>
        </el-card>

        <!-- 源数据预览 -->
        <el-card v-if="rawMaterialData.length > 0 || rawMaterialUnit.length > 0" class="source-card">
          <template #header>
            <span>源数据预览</span>
          </template>
          <el-tabs>
            <el-tab-pane v-if="rawMaterialData.length > 0" label="material_data (原始)">
              <el-table :data="paginatedRawMaterialData" border height="300" style="width: 100%">
                <el-table-column type="index" width="70" label="行号">
                  <template #default="{ $index }">
                    {{ (rawDataPage - 1) * rawDataPageSize + $index + 1 }}
                  </template>
                </el-table-column>
                <el-table-column
                  v-for="col in rawMaterialDataColumns"
                  :key="col.prop"
                  :prop="col.prop"
                  :label="col.label"
                  :min-width="col.minWidth"
                />
              </el-table>
              <div class="source-footer">
                <el-pagination
                  v-model:current-page="rawDataPage"
                  :page-size="rawDataPageSize"
                  :total="rawMaterialDataTotal"
                  layout="total, prev, pager, next"
                  @current-change="handleRawDataPageChange"
                />
                <span class="total-count">共 {{ rawMaterialDataTotal }} 行</span>
              </div>
            </el-tab-pane>
            <el-tab-pane v-if="rawMaterialUnit.length > 0" label="material_unit (原始)">
              <el-table :data="rawMaterialUnit.slice(0, 100)" border height="300" style="width: 100%">
                <el-table-column type="index" width="70" label="行号" />
                <el-table-column
                  v-for="col in rawMaterialUnitColumns"
                  :key="col.prop"
                  :prop="col.prop"
                  :label="col.label"
                  :min-width="col.minWidth"
                />
              </el-table>
              <div class="source-footer">共 {{ rawMaterialUnit.length }} 行</div>
            </el-tab-pane>
          </el-tabs>
        </el-card>
      </el-tab-pane>

      <!-- 合并数据页 -->
      <el-tab-pane label="合并数据" name="merged" :disabled="mergedData.length === 0">
        <el-card class="table-card">
          <template #header>
            <div class="card-header">
              <span class="title">数据预览与编辑</span>
              <div class="header-actions">
                <el-input
                  v-model="searchQuery"
                  placeholder="搜索..."
                  class="search-input"
                  clearable
                />
                <el-button type="success" size="small" @click="addRow">添加</el-button>
                <el-button type="primary" size="small" @click="handleExportExcel">导出Excel</el-button>
                <el-button type="warning" size="small" @click="handleExportEmail">导出邮件</el-button>
              </div>
            </div>
          </template>

          <div class="table-wrapper">
            <el-table
              :data="paginatedData"
              border
              stripe
              height="450"
              style="width: 100%"
            >
              <el-table-column type="index" width="60" label="序号" />

              <el-table-column
                v-for="col in tableColumns"
                :key="col.prop"
                :prop="col.prop"
                :label="col.label"
                :min-width="col.minWidth"
              >
                <template #default="{ row, $index }">
                  <el-input
                    v-model="row[col.prop]"
                    size="small"
                    @change="paginatedData[$index][col.prop] = row[col.prop]"
                  />
                </template>
              </el-table-column>

              <el-table-column label="操作" width="80" fixed="right">
                <template #default="{ $index }">
                  <el-button type="danger" size="small" @click="deleteRow($index)">删除</el-button>
                </template>
              </el-table-column>
            </el-table>
          </div>

          <div class="table-footer">
            <el-pagination
              v-model:current-page="currentPage"
              :page-size="pageSize"
              :total="filteredData.length"
              layout="total, prev, pager, next"
              @current-change="handlePageChange"
            />
            <span class="data-count">共 {{ filteredData.length }} 条数据</span>
          </div>
        </el-card>
      </el-tab-pane>
    </el-tabs>

    <!-- 上传进度对话框 -->
    <el-dialog
      v-model="uploadProgressVisible"
      title="正在处理文件"
      width="90%"
      max-width="500px"
      :close-on-click-modal="false"
      :show-close="false"
    >
      <div class="progress-content">
        <p class="progress-filename">{{ uploadProgress.fileName }}</p>
        <el-progress :percentage="uploadProgress.percent" :stroke-width="15" />
        <div class="progress-info">
          <span>正在处理第 {{ uploadProgress.currentRow }} / {{ uploadProgress.totalRows }} 行</span>
          <span>{{ uploadProgress.percent }}%</span>
        </div>
        <div class="progress-preview" v-if="uploadProgress.currentContent">
          {{ uploadProgress.currentContent }}
        </div>
        <div class="progress-remaining">
          剩余约 {{ Math.max(0, uploadProgress.totalRows - uploadProgress.currentRow) }} 条
        </div>
      </div>
    </el-dialog>

    <!-- 邮件导出对话框 -->
    <el-dialog
      v-model="emailDialogVisible"
      title="导出邮件"
      width="90%"
      max-width="500px"
    >
      <el-form :model="emailForm" label-width="80px">
        <el-form-item label="收件人">
          <el-input v-model="emailForm.to" placeholder="请输入收件人邮箱" />
        </el-form-item>
        <el-form-item label="主题">
          <el-input v-model="emailForm.subject" placeholder="邮件主题" />
        </el-form-item>
      </el-form>
      <div class="email-preview">
        <h4>邮件预览（前20条）</h4>
        <div class="preview-html-wrapper" v-html="emailPreviewHtml"></div>
        <p class="preview-footer">附件将包含全部 {{ mergedData.length }} 条数据</p>
      </div>
      <template #footer>
        <el-button @click="emailDialogVisible = false">取消</el-button>
        <el-button type="primary" @click="confirmExportEmail">确认导出</el-button>
      </template>
    </el-dialog>
  </div>
</template>

<style scoped>
.container {
  padding: 12px;
  min-height: 100vh;
  background: #f5f5f5;
}

.main-tabs {
  background: white;
  padding: 16px;
  border-radius: 4px;
}

.upload-card {
  margin-bottom: 16px;
}

.card-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  flex-wrap: wrap;
  gap: 10px;
}

.card-header .title {
  font-weight: bold;
  font-size: 16px;
}

.header-actions {
  display: flex;
  align-items: center;
  flex-wrap: wrap;
  gap: 8px;
}

.search-input {
  width: 120px !important;
}

.upload-section {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
  gap: 20px;
}

.upload-item h4 {
  margin-bottom: 5px;
  color: #409eff;
  font-size: 14px;
  display: flex;
  align-items: center;
  gap: 8px;
}

.source-tag {
  font-size: 11px;
  color: #909399;
  font-weight: normal;
  background: #f0f0f0;
  padding: 2px 6px;
  border-radius: 3px;
}

.upload-hint {
  font-size: 12px;
  color: #909399;
  margin-bottom: 10px;
}

.upload-status {
  margin-top: 16px;
  display: flex;
  gap: 10px;
  flex-wrap: wrap;
}

.calculate-btn-wrapper {
  margin-top: 20px;
  text-align: center;
}

.source-card {
  margin-top: 16px;
}

.source-footer {
  margin-top: 10px;
  display: flex;
  justify-content: space-between;
  align-items: center;
}

.source-footer .total-count {
  color: #606266;
  font-size: 13px;
}

.table-card {
  margin-bottom: 12px;
}

.table-wrapper {
  overflow-x: auto;
}

.table-footer {
  margin-top: 10px;
  display: flex;
  justify-content: space-between;
  align-items: center;
}

.data-count {
  color: #606266;
  font-size: 13px;
}

/* 进度对话框 */
.progress-content {
  padding: 10px 0;
}

.progress-filename {
  font-weight: bold;
  margin-bottom: 15px;
  color: #409eff;
}

.progress-info {
  display: flex;
  justify-content: space-between;
  margin-top: 10px;
  color: #606266;
  font-size: 13px;
}

.progress-preview {
  margin-top: 10px;
  padding: 8px;
  background: #f5f5f5;
  border-radius: 4px;
  font-size: 12px;
  color: #909399;
  max-height: 60px;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.progress-remaining {
  margin-top: 10px;
  color: #909399;
  font-size: 12px;
}

/* 邮件预览 */
.email-preview {
  margin-top: 15px;
  border-top: 1px solid #eee;
  padding-top: 15px;
}

.email-preview h4 {
  margin-bottom: 10px;
  font-size: 14px;
}

.preview-html-wrapper {
  max-height: 300px;
  overflow: auto;
  border: 1px solid #ddd;
  border-radius: 4px;
  padding: 10px;
  background: #fafafa;
}

.preview-footer {
  margin-top: 10px;
  color: #909399;
  font-size: 12px;
}

@media (max-width: 768px) {
  .container {
    padding: 8px;
  }

  .card-header {
    flex-direction: column;
    align-items: flex-start;
  }

  .header-actions {
    width: 100%;
  }

  .search-input {
    width: 100% !important;
  }
}
</style>
