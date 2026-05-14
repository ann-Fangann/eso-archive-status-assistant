<template>
  <div class="drawing-matcher-container">
    <el-card shadow="hover" class="drawing-card">
      <template #header>
        <div class="card-header">
          <span>图纸未完成清单</span>
        </div>
      </template>

      <div class="upload-row">
        <div class="upload-column">
          <h4>图纸/数模主清单</h4>
          <el-upload
            drag
            :auto-upload="false"
            :on-change="(file: UploadFile) => handleChange(file, 'club')"
            accept=".xlsx,.xls"
          >
            <el-icon class="el-icon--upload"><upload-filled /></el-icon>
            <div class="el-upload__text">
              将图纸主清单拖到此处，或 <em>点击上传</em>
            </div>
            <template #tip>
              <div class="el-upload__tip">需包含零件号、数模计划/实际日期、图纸实际日期等字段</div>
            </template>
          </el-upload>
        </div>

        <div class="upload-column">
          <h4>图纸发布/归档清单（可选）</h4>
          <el-upload
            drag
            :auto-upload="false"
            :on-change="(file: UploadFile) => handleChange(file, 'published')"
            accept=".xlsx,.xls"
          >
            <el-icon class="el-icon--upload"><upload-filled /></el-icon>
            <div class="el-upload__text">
              将图纸发布清单拖到此处，或 <em>点击上传</em>
            </div>
            <template #tip>
              <div class="el-upload__tip">可用于补充图纸实际发布日期</div>
            </template>
          </el-upload>
        </div>
      </div>

      <div class="date-section">
        <span class="date-label">未完成统计日期</span>
        <el-date-picker
          v-model="targetDate"
          type="date"
          value-format="YYYY-MM-DD"
          placeholder="选择统计日期"
          style="width: 180px;"
        />
      </div>

      <el-button
        type="success"
        size="large"
        style="width: 100%;"
        :disabled="!clubFile"
        :loading="loading"
        @click="handleUpload"
      >
        {{ loading ? '正在生成...' : '生成图纸未完成清单' }}
      </el-button>

      <el-alert
        v-if="message"
        :title="message"
        :type="messageType"
        show-icon
        :closable="false"
        style="margin-top: 20px;"
      />

      <div v-if="result" style="margin-top: 20px;">
        <el-divider />
        <h3>汇总结果</h3>
        <el-descriptions :column="2" border>
          <el-descriptions-item label="统计日期">{{ summary.target_date }}</el-descriptions-item>
          <el-descriptions-item label="未完成数量">{{ summary.unfinished_count }}</el-descriptions-item>
          <el-descriptions-item label="排除D行数量">{{ summary.delete_excluded_count }}</el-descriptions-item>
          <el-descriptions-item label="统计口径" :span="2">{{ summary.formula }}</el-descriptions-item>
        </el-descriptions>

        <div class="section-header">
          <h3>未完成清单</h3>
          <div class="section-actions">
            <span>共 {{ unfinishedRows.length }} 条</span>
            <el-button size="small" type="primary" :disabled="!unfinishedRows.length" @click="exportDrawingList">
              导出图纸未完成清单
            </el-button>
          </div>
        </div>

        <el-table :data="unfinishedRows" style="width: 100%" stripe max-height="460">
          <el-table-column
            v-for="col in displayColumns"
            :key="col"
            :prop="col"
            :label="col"
            :min-width="getColumnWidth(col)"
            show-overflow-tooltip
          />
        </el-table>
      </div>
    </el-card>
  </div>
</template>

<script setup lang="ts">
import { computed, ref } from 'vue'
import axios from 'axios'
import { UploadFilled } from '@element-plus/icons-vue'
import { ElMessage } from 'element-plus'
import type { UploadFile } from 'element-plus'
import * as XLSX from 'xlsx-js-style'

const clubFile = ref<File | null>(null)
const publishedFile = ref<File | null>(null)
const loading = ref(false)
const message = ref('')
const messageType = ref<'success' | 'error' | 'warning' | 'info'>('info')
const result = ref<any>(null)

const getYesterday = () => {
  const day = new Date()
  day.setDate(day.getDate() - 1)
  const year = day.getFullYear()
  const month = String(day.getMonth() + 1).padStart(2, '0')
  const date = String(day.getDate()).padStart(2, '0')
  return `${year}-${month}-${date}`
}

const targetDate = ref(getYesterday())
const summary = computed(() => result.value?.select_result?.summary || {})
const unfinishedRows = computed(() => result.value?.select_result?.data || [])
const groupStats = computed(() => result.value?.select_result?.group_stats || [])
const typeStats = computed(() => result.value?.select_result?.type_stats || [])

const preferredColumns = [
  '操作类型',
  '产品',
  '车型年',
  '零件号',
  '功能组',
  '工程师',
  '图纸号',
  '未完成类型',
  '数模状态',
  '图纸状态',
  '图纸要求完成日期',
  '图纸实际日期'
]

const displayColumns = computed(() => {
  const keySet = new Set<string>()
  unfinishedRows.value.forEach((row: Record<string, any>) => {
    Object.keys(row).forEach(key => keySet.add(key))
  })
  return [
    ...preferredColumns.filter(key => keySet.has(key)),
    ...Array.from(keySet).filter(key => !preferredColumns.includes(key))
  ]
})

const handleChange = (uploadFile: UploadFile, type: string) => {
  if (type === 'club') {
    clubFile.value = uploadFile.raw || null
  } else {
    publishedFile.value = uploadFile.raw || null
  }
  result.value = null
  message.value = ''
}

const handleUpload = async () => {
  if (!clubFile.value) {
    message.value = '请先上传图纸/数模主清单'
    messageType.value = 'warning'
    return
  }

  const formData = new FormData()
  formData.append('club_file', clubFile.value)
  if (publishedFile.value) {
    formData.append('published_file', publishedFile.value)
  }
  if (targetDate.value) {
    formData.append('target_date', targetDate.value)
  }

  try {
    loading.value = true
    message.value = '正在生成图纸未完成清单...'
    messageType.value = 'info'
    const response = await axios.post('/api/upload-drawing-excel/', formData, {
      headers: { 'Content-Type': 'multipart/form-data' },
      timeout: 300000
    })
    result.value = response.data
    message.value = '图纸未完成清单已生成'
    messageType.value = 'success'
  } catch (error: any) {
    console.error('图纸未完成清单生成失败:', error)
    const detail = error.response?.data?.detail
    message.value = typeof detail === 'object'
      ? `${detail.message}。当前识别字段：${JSON.stringify(detail.detected_columns)}`
      : `生成失败: ${detail || error.message || '未知错误'}`
    messageType.value = 'error'
    result.value = null
  } finally {
    loading.value = false
  }
}

const getColumnWidth = (prop: string) => {
  if (['零件号', '工程师', '图纸号', '图纸要求完成日期'].includes(prop)) return 160
  if (prop.includes('日期') || prop.includes('Date')) return 140
  if (['未完成类型', '数模状态', '图纸状态'].includes(prop)) return 150
  return 120
}

const getExportDateString = () => String(summary.value?.target_date || targetDate.value || getYesterday()).replace(/-/g, '')

const styleWorksheet = (ws: XLSX.WorkSheet, headerColor = '4472C4') => {
  const range = ws['!ref'] ? XLSX.utils.decode_range(ws['!ref']) : null
  if (!range) return

  for (let row = range.s.r; row <= range.e.r; row++) {
    for (let col = range.s.c; col <= range.e.c; col++) {
      const address = XLSX.utils.encode_cell({ r: row, c: col })
      const cell = ws[address]
      if (!cell) continue

      const isHeader = row === 0
      const isBlankBodyCell = !isHeader && (cell.v === undefined || cell.v === null || cell.v === '')
      if (isBlankBodyCell) continue

      cell.s = {
        border: {
          top: { style: 'thin', color: { rgb: 'D9E2F3' } },
          bottom: { style: 'thin', color: { rgb: 'D9E2F3' } },
          left: { style: 'thin', color: { rgb: 'D9E2F3' } },
          right: { style: 'thin', color: { rgb: 'D9E2F3' } }
        },
        alignment: { vertical: 'center', wrapText: true }
      }

      if (isHeader) {
        cell.s.fill = { fgColor: { rgb: headerColor } }
        cell.s.font = { color: { rgb: 'FFFFFF' }, bold: true }
      }
    }
  }
}

const appendJsonSheet = (wb: XLSX.WorkBook, rows: Record<string, any>[], sheetName: string) => {
  const ws = XLSX.utils.json_to_sheet(rows.length ? rows : [{}])
  styleWorksheet(ws)
  XLSX.utils.book_append_sheet(wb, ws, sheetName)
}

const exportDrawingList = () => {
  if (!unfinishedRows.value.length) {
    ElMessage.warning('当前没有可导出的图纸未完成清单')
    return
  }

  const rows = unfinishedRows.value.map((row: Record<string, any>) => {
    const output: Record<string, any> = {}
    displayColumns.value.forEach(col => {
      output[col] = row[col] ?? ''
    })
    return output
  })

  const wb = XLSX.utils.book_new()
  const detailWs = XLSX.utils.json_to_sheet(rows)
  styleWorksheet(detailWs)
  detailWs['!cols'] = displayColumns.value.map(col => ({ wch: Math.min(Math.max(Math.round(getColumnWidth(col) / 8), 12), 30) }))
  XLSX.utils.book_append_sheet(wb, detailWs, '图纸未完成清单')

  appendJsonSheet(wb, [
    { 指标: '统计日期', 数值: summary.value.target_date || '' },
    { 指标: '未完成数量', 数值: summary.value.unfinished_count || 0 },
    { 指标: '排除D行数量', 数值: summary.value.delete_excluded_count || 0 },
    { 指标: '统计口径', 数值: summary.value.formula || '' },
  ], '汇总结果')
  appendJsonSheet(wb, groupStats.value.map((item: any) => ({ 功能组: item['功能组'], 未完成数量: item.count })), '部门统计')
  appendJsonSheet(wb, typeStats.value.map((item: any) => ({ 未完成类型: item['未完成类型'], 数量: item.count })), '类型统计')

  XLSX.writeFile(wb, `图纸未完成清单-${getExportDateString()}.xlsx`)
  ElMessage.success('图纸未完成清单已导出')
}
</script>

<style scoped>
.drawing-matcher-container {
  height: 100%;
  display: flex;
  flex-direction: column;
}

.drawing-card {
  flex: 1;
}

.card-header {
  font-size: 18px;
  font-weight: bold;
}

.upload-row {
  display: flex;
  gap: 20px;
  margin-bottom: 20px;
}

.upload-column {
  flex: 1;
}

.date-section {
  display: flex;
  align-items: center;
  gap: 12px;
  margin-bottom: 20px;
}

.date-label {
  color: #303133;
  font-weight: 600;
}

.section-header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  margin: 20px 0 12px;
}

.section-header h3 {
  margin: 0;
}

.section-actions {
  display: flex;
  align-items: center;
  gap: 10px;
}

.section-actions span {
  color: #606266;
  font-size: 14px;
}

@media (max-width: 768px) {
  .upload-row {
    flex-direction: column;
  }
}
</style>
