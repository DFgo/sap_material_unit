<script setup>
import { ref, computed } from 'vue'
import { ElUpload, ElButton, ElTable, ElTableColumn, ElInput, ElMessage, ElCard, ElDialog } from 'element-plus'
import {
  readExcelFile,
  parseMaterialData,
  parseMaterialUnit,
  processMerge,
  exportToExcel,
  generateEmailAttachment,
  getUnitGroupHeaders
} from '../utils/excelProcessor'

// 状态
const materialDataFile = ref(null)
const materialUnitFile = ref(null)
const materialData = ref([])
const materialUnit = ref([])
const mergedData = ref([])
const maxUnits = ref(0)
const tableHeaders = ref([])
const isProcessing = ref(false)
const searchQuery = ref('')
const emailDialogVisible = ref(false)
const emailContent = ref('')

// 过滤后的数据
const filteredData = computed(() => {
  if (!searchQuery.value) return mergedData.value
  const q = searchQuery.value.toLowerCase()
  return mergedData.value.filter(row =>
    row.物料?.toString().toLowerCase().includes(q) ||
    row.物料描述?.toString().toLowerCase().includes(q)
  )
})

// 表格列（使用 min-width 自适应）
const tableColumns = computed(() => {
  const cols = [
    { prop: '物料', label: '物料', minWidth: 100 },
    { prop: '物料描述', label: '物料描述', minWidth: 150 },
    { prop: '基本计量单位', label: '基本计量单位', minWidth: 90 }
  ]

  for (let i = 1; i <= maxUnits.value; i++) {
    cols.push(
      { prop: `分母${i}`, label: `分母${i}`, minWidth: 70 },
      { prop: `基本计量单位${i}`, label: `基本单位${i}`, minWidth: 80 },
      { prop: `可选单位${i}`, label: `可选单位${i}`, minWidth: 80 },
      { prop: `等于${i}`, label: `=`, minWidth: 50 },
      { prop: `计数器${i}`, label: `计数器${i}`, minWidth: 70 }
    )
  }
  return cols
})

// 文件上传处理
async function handleMaterialDataUpload(file) {
  materialDataFile.value = file.raw
  try {
    const result = await readExcelFile(file.raw)
    materialData.value = parseMaterialData(result)
    ElMessage.success(`material_data.xlsx 读取成功，共 ${materialData.value.length} 条数据`)
    tryMerge()
  } catch (err) {
    ElMessage.error('读取 material_data.xlsx 失败')
    console.error(err)
  }
  return false
}

async function handleMaterialUnitUpload(file) {
  materialUnitFile.value = file.raw
  try {
    const result = await readExcelFile(file.raw)
    materialUnit.value = parseMaterialUnit(result)
    ElMessage.success(`material_unit.xlsx 读取成功，共 ${materialUnit.value.length} 条数据`)
    tryMerge()
  } catch (err) {
    ElMessage.error('读取 material_unit.xlsx 失败')
    console.error(err)
  }
  return false
}

// 合并数据
function tryMerge() {
  if (materialData.value.length === 0 || materialUnit.value.length === 0) {
    return
  }

  isProcessing.value = true
  try {
    const result = processMerge(materialData.value, materialUnit.value)
    mergedData.value = result.mergedData
    maxUnits.value = result.maxUnits
    tableHeaders.value = result.headers
    ElMessage.success(`合并完成！共 ${mergedData.value.length} 条数据，最大 ${maxUnits.value} 个单位转换`)
  } catch (err) {
    ElMessage.error('合并失败')
    console.error(err)
  } finally {
    isProcessing.value = false
  }
}

// 导出 Excel
function handleExportExcel() {
  if (mergedData.value.length === 0) {
    ElMessage.warning('没有数据可导出')
    return
  }
  try {
    exportToExcel(mergedData.value, maxUnits.value, 'merged_output.xlsx')
    ElMessage.success('Excel 导出成功')
  } catch (err) {
    ElMessage.error('导出失败')
    console.error(err)
  }
}

// 生成邮件
function handleGenerateEmail() {
  if (mergedData.value.length === 0) {
    ElMessage.warning('没有数据可生成邮件')
    return
  }
  try {
    const base64 = generateEmailAttachment(mergedData.value, maxUnits.value)
    emailContent.value = base64
    emailDialogVisible.value = true
  } catch (err) {
    ElMessage.error('生成邮件失败')
    console.error(err)
  }
}

// 下载附件
function downloadAttachment() {
  const binary = atob(emailContent.value)
  const array = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) {
    array[i] = binary.charCodeAt(i)
  }
  const blob = new Blob([array], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = 'material_data.xlsx'
  a.click()
  URL.revokeObjectURL(url)
  ElMessage.success('附件已生成，点击复制 Base64 内容')
}

// 复制 Base64
function copyBase64() {
  navigator.clipboard.writeText(emailContent.value)
  ElMessage.success('Base64 内容已复制到剪贴板')
}

// 删除行
function deleteRow(index) {
  mergedData.value.splice(index, 1)
  ElMessage.success('已删除')
}

// 添加行
function addRow() {
  mergedData.value.push({
    物料: '',
    物料描述: '',
    基本计量单位: ''
  })
}

// 行样式
function getRowClass({ rowIndex }) {
  if (rowIndex === 0) return 'header-row'
  return ''
}
</script>

<template>
  <div class="container">
    <el-card class="upload-card">
      <template #header>
        <div class="card-header">
          <span>上传 Excel 文件</span>
        </div>
      </template>

      <div class="upload-section">
        <div class="upload-item">
          <h4>material_data.xlsx</h4>
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
          <h4>material_unit.xlsx</h4>
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
    </el-card>

    <el-card v-if="mergedData.length > 0" class="table-card">
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
            <el-button type="primary" size="small" @click="handleExportExcel">导出</el-button>
            <el-button type="warning" size="small" @click="handleGenerateEmail">邮件</el-button>
          </div>
        </div>
      </template>

      <div class="table-wrapper">
        <el-table
          :data="filteredData"
          border
          stripe
          height="500"
          :row-class-name="getRowClass"
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
                @change="mergedData[$index][col.prop] = row[col.prop]"
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
        <span>共 {{ filteredData.length }} 条数据</span>
      </div>
    </el-card>

    <el-card v-else class="empty-card">
      <div class="empty-state">
        <p>请上传两个 Excel 文件开始处理</p>
      </div>
    </el-card>

    <!-- 邮件对话框 -->
    <el-dialog
      v-model="emailDialogVisible"
      title="邮件附件"
      width="90%"
      max-width="600px"
    >
      <div class="email-dialog-content">
        <p>附件 Base64 内容已生成，可以下载或复制使用：</p>
        <el-input
          v-model="emailContent"
          type="textarea"
          :rows="6"
          style="margin-top: 10px"
        />
      </div>
      <template #footer>
        <el-button @click="copyBase64">复制 Base64</el-button>
        <el-button type="primary" @click="downloadAttachment">下载附件</el-button>
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

.upload-card {
  margin-bottom: 12px;
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
}

.upload-hint {
  font-size: 12px;
  color: #909399;
  margin-bottom: 10px;
}

.table-card {
  margin-bottom: 12px;
}

.table-wrapper {
  overflow-x: auto;
}

.table-footer {
  margin-top: 10px;
  color: #606266;
  font-size: 13px;
}

.empty-card {
  text-align: center;
}

.empty-state {
  padding: 60px 0;
  color: #909399;
}

.email-dialog-content {
  padding: 10px 0;
  word-break: break-all;
}

.email-dialog-content .el-textarea {
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
