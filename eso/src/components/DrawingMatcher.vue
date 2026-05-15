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
              <div class="el-upload__tip">可用于补充并回填图纸实际发布日期</div>
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
        <el-button
          type="primary"
          plain
          :disabled="!result"
          :loading="refreshing"
          @click="refreshDrawingList"
        >
          {{ refreshing ? '正在生成...' : '重新生成未完成清单' }}
        </el-button>
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
          <el-descriptions-item label="数模到期数量">{{ summary.model_due_count || 0 }}</el-descriptions-item>
          <el-descriptions-item label="数模已完成数量">{{ summary.model_completed_count || 0 }}</el-descriptions-item>
          <el-descriptions-item label="图纸到期数量">{{ summary.drawing_due_count || 0 }}</el-descriptions-item>
          <el-descriptions-item label="图纸已完成数量">{{ summary.drawing_completed_count || 0 }}</el-descriptions-item>
          <el-descriptions-item label="本次回填数量">{{ summary.matched_completed_count ?? 0 }}</el-descriptions-item>
          <el-descriptions-item label="排除D行数量">{{ summary.delete_excluded_count || 0 }}</el-descriptions-item>
          <el-descriptions-item label="统计口径" :span="2">{{ summary.formula }}</el-descriptions-item>
        </el-descriptions>

        <el-alert
          v-if="result.persistence"
          :title="result.persistence.message"
          :type="result.persistence.saved ? 'success' : 'warning'"
          show-icon
          :closable="false"
          style="margin-top: 16px;"
        />

        <div v-if="result.modified_workbook" style="margin-top: 16px;">
          <el-alert
            v-if="result.modified_workbook.error"
            :title="result.modified_workbook.error"
            type="warning"
            show-icon
            :closable="false"
          />
          <el-button
            v-else
            type="primary"
            @click="downloadFile(result.modified_workbook.download_url)"
          >
            下载已回填的图纸主清单
          </el-button>
        </div>

        <div class="report-panel">
          <div class="section-header">
            <h4>邮件汇报文本</h4>
            <el-button size="small" plain @click="copyReportText">复制文本</el-button>
          </div>
          <el-input
            :model-value="reportText"
            type="textarea"
            :autosize="{ minRows: 4, maxRows: 8 }"
            readonly
          />
        </div>

        <div class="section-header">
          <h3>未完成清单</h3>
          <div class="section-actions">
            <span>共 {{ unfinishedRows.length }} 条</span>
            <el-popover
              placement="bottom"
              title="选择显示和导出的列"
              width="320"
              trigger="click"
            >
              <div class="column-options">
                <el-checkbox
                  :model-value="isAllColumnsSelected"
                  :indeterminate="isSomeColumnsSelected"
                  @change="toggleColumns"
                >
                  全选/全不选
                </el-checkbox>
                <el-checkbox-group
                  v-model="selectedColumns"
                  class="column-checkbox-group"
                  @change="saveColumnSelection"
                >
                  <el-checkbox
                    v-for="col in allColumns"
                    :key="col.prop"
                    :label="col.prop"
                    class="column-checkbox"
                  >
                    {{ col.label }}
                  </el-checkbox>
                </el-checkbox-group>
              </div>
              <template #reference>
                <el-button size="small" plain>选择列</el-button>
              </template>
            </el-popover>
            <el-button
              size="small"
              type="primary"
              :disabled="!unfinishedRows.length || !selectedColumns.length"
              @click="exportDrawingList"
            >
              导出图纸未完成清单
            </el-button>
          </div>
        </div>

        <el-table :data="unfinishedRows" style="width: 100%" stripe max-height="460">
          <el-table-column
            v-for="col in selectedColumnDefs"
            :key="col.prop"
            :prop="col.prop"
            :label="col.label"
            :min-width="col.width"
            show-overflow-tooltip
          />
        </el-table>

        <div class="charts-section">
          <el-divider />
          <h3>数据可视化分析</h3>
          <div class="chart-container">
            <h4>功能组图纸未完成分布</h4>
            <div ref="barChartRef" class="chart-wrapper"></div>
          </div>
          <div class="chart-container">
            <h4>图纸/数模未完成类型分布</h4>
            <div ref="pieChartRef" class="chart-wrapper"></div>
          </div>
        </div>
      </div>
    </el-card>
  </div>
</template>

<script setup lang="ts">
import { computed, nextTick, onUnmounted, ref, watch } from 'vue'
import axios from 'axios'
import { UploadFilled } from '@element-plus/icons-vue'
import { ElMessage } from 'element-plus'
import type { UploadFile } from 'element-plus'
import * as echarts from 'echarts'
import * as XLSX from 'xlsx-js-style'

const clubFile = ref<File | null>(null)
const publishedFile = ref<File | null>(null)
const loading = ref(false)
const refreshing = ref(false)
const message = ref('')
const messageType = ref<'success' | 'error' | 'warning' | 'info'>('info')
const result = ref<any>(null)
const selectedColumns = ref<string[]>([])
const chartDataReady = ref(false)
const barChartRef = ref<HTMLElement | null>(null)
const pieChartRef = ref<HTMLElement | null>(null)

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

const defaultColumns = [
  '操作类型',
  '产品',
  '车型年',
  '零件号',
  '功能组',
  '工程师',
  '未完成类型',
  '图纸要求完成日期',
  '图纸实际日期'
]

const reportText = computed(() => {
  if (!summary.value?.target_date) return ''
  const topGroups = groupStats.value
    .slice(0, 5)
    .map((item: any) => `${item['功能组'] || '未知功能组'} ${item.count || 0} 项`)
    .join('，')

  return [
    `截至 ${summary.value.target_date}，图纸/数模未完成共 ${summary.value.unfinished_count || 0} 项，其中数模到期 ${summary.value.model_due_count || 0} 项，图纸到期 ${summary.value.drawing_due_count || 0} 项。`,
    `本次根据图纸发布/归档清单回填图纸实际日期 ${summary.value.matched_completed_count ?? 0} 项；未完成统计已排除操作类型为 D 的记录 ${summary.value.delete_excluded_count || 0} 项。`,
    topGroups ? `未完成主要分布：${topGroups}。` : '当前统计日期没有图纸/数模未完成项。'
  ].join('\n')
})

const allColumns = computed(() => {
  const keySet = new Set<string>()
  unfinishedRows.value.forEach((row: Record<string, any>) => {
    Object.keys(row).forEach(key => keySet.add(key))
  })
  const orderedKeys = [
    ...preferredColumns.filter(key => keySet.has(key)),
    ...Array.from(keySet).filter(key => !preferredColumns.includes(key))
  ]
  return orderedKeys.map(prop => ({
    prop,
    label: prop,
    width: getColumnWidth(prop)
  }))
})

const selectedColumnDefs = computed(() => {
  const selectedSet = new Set(selectedColumns.value)
  return allColumns.value.filter(col => selectedSet.has(col.prop))
})

const isAllColumnsSelected = computed(() => allColumns.value.length > 0 && selectedColumns.value.length === allColumns.value.length)
const isSomeColumnsSelected = computed(() => selectedColumns.value.length > 0 && selectedColumns.value.length < allColumns.value.length)

const getSavedColumnSelection = () => {
  try {
    const saved = localStorage.getItem('drawing_unfinished_columns')
    return saved ? JSON.parse(saved) : []
  } catch (error) {
    console.error('读取图纸列选择失败:', error)
    return []
  }
}

const saveColumnSelection = () => {
  localStorage.setItem('drawing_unfinished_columns', JSON.stringify(selectedColumns.value))
}

const resetColumnSelection = () => {
  const allKeys = allColumns.value.map(col => col.prop)
  if (!allKeys.length) {
    selectedColumns.value = []
    return
  }
  const saved = getSavedColumnSelection().filter((key: string) => allKeys.includes(key))
  const defaults = defaultColumns.filter(key => allKeys.includes(key))
  selectedColumns.value = saved.length ? saved : defaults
}

const toggleColumns = (checked: boolean) => {
  selectedColumns.value = checked ? allColumns.value.map(col => col.prop) : []
  saveColumnSelection()
}

watch(allColumns, (columns) => {
  const allKeys = columns.map(col => col.prop)
  const validSelected = selectedColumns.value.filter(key => allKeys.includes(key))
  if (!allKeys.length) {
    selectedColumns.value = []
    return
  }
  if (!validSelected.length) {
    resetColumnSelection()
    return
  }
  if (validSelected.length !== selectedColumns.value.length) {
    selectedColumns.value = validSelected
    saveColumnSelection()
  }
}, { immediate: true })

const handleChange = (uploadFile: UploadFile, type: string) => {
  if (type === 'club') {
    clubFile.value = uploadFile.raw || null
  } else {
    publishedFile.value = uploadFile.raw || null
  }
  result.value = null
  message.value = ''
  chartDataReady.value = false
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
    chartDataReady.value = false
    const response = await axios.post('/api/upload-drawing-excel/', formData, {
      headers: { 'Content-Type': 'multipart/form-data' },
      timeout: 300000
    })
    result.value = response.data
    message.value = '图纸未完成清单已生成'
    messageType.value = 'success'
    await nextTick()
    prepareCharts()
  } catch (error: any) {
    console.error('图纸未完成清单生成失败:', error)
    const detail = error.response?.data?.detail
    message.value = typeof detail === 'object'
      ? `${detail.message}。当前识别字段：${JSON.stringify(detail.detected_columns)}`
      : `生成失败: ${detail || error.message || '未知错误'}`
    messageType.value = 'error'
    result.value = null
    chartDataReady.value = false
  } finally {
    loading.value = false
  }
}

const refreshDrawingList = async () => {
  if (!result.value) {
    message.value = '请先上传并处理图纸文件'
    messageType.value = 'warning'
    return
  }

  try {
    refreshing.value = true
    message.value = '正在按统计日期生成图纸未完成清单...'
    messageType.value = 'info'
    chartDataReady.value = false
    const response = await axios.post('/api/drawing-unfinished-list/', {
      target_date: targetDate.value
    }, {
      timeout: 120000
    })
    result.value.select_result = response.data.select_result
    if (response.data.modified_workbook) result.value.modified_workbook = response.data.modified_workbook
    if (response.data.persistence) result.value.persistence = response.data.persistence
    message.value = '图纸未完成清单已重新生成'
    messageType.value = 'success'
    await nextTick()
    prepareCharts()
  } catch (error: any) {
    console.error('重新生成图纸未完成清单失败:', error)
    message.value = '重新生成失败: ' + (error.response?.data?.detail || error.message || '未知错误')
    messageType.value = 'error'
  } finally {
    refreshing.value = false
  }
}

const downloadFile = (downloadUrl: string) => {
  window.open(`/api${downloadUrl}`, '_blank')
}

const copyReportText = async () => {
  if (!reportText.value) {
    ElMessage.warning('暂无可复制的汇报文本')
    return
  }
  try {
    await navigator.clipboard.writeText(reportText.value)
    ElMessage.success('汇报文本已复制')
  } catch (error) {
    console.error('复制汇报文本失败:', error)
    ElMessage.warning('复制失败，请手动选择文本复制')
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
  if (!selectedColumns.value.length) {
    ElMessage.warning('请至少选择一列')
    return
  }

  const rows = unfinishedRows.value.map((row: Record<string, any>) => {
    const output: Record<string, any> = {}
    selectedColumnDefs.value.forEach(col => {
      output[col.label] = row[col.prop] ?? ''
    })
    return output
  })

  const wb = XLSX.utils.book_new()
  const detailWs = XLSX.utils.json_to_sheet(rows)
  styleWorksheet(detailWs)
  detailWs['!cols'] = selectedColumnDefs.value.map(col => ({ wch: Math.min(Math.max(Math.round((col.width || 120) / 8), 12), 30) }))
  XLSX.utils.book_append_sheet(wb, detailWs, '图纸未完成清单')

  appendJsonSheet(wb, [
    { 指标: '统计日期', 数值: summary.value.target_date || '' },
    { 指标: '未完成数量', 数值: summary.value.unfinished_count || 0 },
    { 指标: '数模到期数量', 数值: summary.value.model_due_count || 0 },
    { 指标: '数模已完成数量', 数值: summary.value.model_completed_count || 0 },
    { 指标: '图纸到期数量', 数值: summary.value.drawing_due_count || 0 },
    { 指标: '图纸已完成数量', 数值: summary.value.drawing_completed_count || 0 },
    { 指标: '本次回填数量', 数值: summary.value.matched_completed_count ?? 0 },
    { 指标: '排除D行数量', 数值: summary.value.delete_excluded_count || 0 },
    { 指标: '统计口径', 数值: summary.value.formula || '' },
  ], '汇总结果')
  appendJsonSheet(wb, groupStats.value.map((item: any) => ({ 功能组: item['功能组'], 未完成数量: item.count })), '部门统计')
  appendJsonSheet(wb, typeStats.value.map((item: any) => ({ 未完成类型: item['未完成类型'], 数量: item.count })), '类型统计')

  XLSX.writeFile(wb, `图纸未完成清单-${getExportDateString()}.xlsx`)
  ElMessage.success('图纸未完成清单已导出')
}

const prepareCharts = async (retryCount = 0) => {
  chartDataReady.value = true
  await nextTick()
  window.requestAnimationFrame(() => {
    if (!barChartRef.value || !pieChartRef.value) {
      if (retryCount < 5) {
        window.setTimeout(() => prepareCharts(retryCount + 1), 80)
      }
      return
    }
    generateBarChart()
    generatePieChart()
  })
}

const generateBarChart = () => {
  if (!barChartRef.value) {
    console.warn('图纸柱状图容器尚未渲染')
    return
  }
  const existingChart = echarts.getInstanceByDom(barChartRef.value)
  if (existingChart) existingChart.dispose()
  const chart = echarts.init(barChartRef.value)

  const data = groupStats.value.reduce((acc: Record<string, number>, item: any) => {
    acc[item['功能组'] || '未知功能组'] = Number(item.count || 0)
    return acc
  }, {})

  if (!Object.keys(data).length) {
    chart.setOption({
      title: { text: '该统计日期没有未完成项', left: 'center', top: 'center' }
    })
    return
  }

  chart.setOption({
    title: { text: '各功能组未完成数量', left: 'center' },
    tooltip: { trigger: 'axis', axisPointer: { type: 'shadow' } },
    xAxis: {
      type: 'category',
      data: Object.keys(data),
      axisLabel: { interval: 0, rotate: 45, fontSize: 12 }
    },
    yAxis: { type: 'value', name: '数量' },
    series: [{
      data: Object.values(data),
      type: 'bar',
      itemStyle: { color: '#409eff' }
    }]
  })
}

const generatePieChart = () => {
  if (!pieChartRef.value) {
    console.warn('图纸饼图容器尚未渲染')
    return
  }
  const existingChart = echarts.getInstanceByDom(pieChartRef.value)
  if (existingChart) existingChart.dispose()
  const chart = echarts.init(pieChartRef.value)

  const data = typeStats.value.map((item: any) => ({
    name: item['未完成类型'] || '未知类型',
    value: Number(item.count || 0)
  }))

  if (!data.length) {
    chart.setOption({
      title: { text: '没有可显示的数据', left: 'center', top: 'center' }
    })
    return
  }

  chart.setOption({
    title: { text: '未完成类型占比', left: 'center' },
    tooltip: { trigger: 'item', formatter: '{a} <br/>{b}: {c} ({d}%)' },
    legend: { orient: 'horizontal', bottom: 'bottom' },
    series: [{
      name: '未完成类型',
      type: 'pie',
      radius: ['40%', '70%'],
      itemStyle: { borderRadius: 8, borderColor: '#fff', borderWidth: 2 },
      label: { show: true, formatter: '{b}: {c} ({d}%)' },
      data
    }]
  })
}

const resizeCharts = () => {
  if (barChartRef.value) echarts.getInstanceByDom(barChartRef.value)?.resize()
  if (pieChartRef.value) echarts.getInstanceByDom(pieChartRef.value)?.resize()
}

window.addEventListener('resize', resizeCharts)

watch(() => result.value?.select_result, () => {
  if (result.value) {
    prepareCharts()
  }
}, { deep: true })

onUnmounted(() => {
  window.removeEventListener('resize', resizeCharts)
  if (barChartRef.value) echarts.getInstanceByDom(barChartRef.value)?.dispose()
  if (pieChartRef.value) echarts.getInstanceByDom(pieChartRef.value)?.dispose()
})
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
  min-width: 0;
}

.upload-column h4 {
  margin-top: 0;
  margin-bottom: 15px;
  text-align: center;
  color: #409eff;
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

.section-header h3,
.section-header h4 {
  margin: 0;
}

.section-actions {
  display: flex;
  align-items: center;
  gap: 10px;
  flex-wrap: wrap;
  justify-content: flex-end;
}

.section-actions span {
  color: #606266;
  font-size: 14px;
}

.report-panel {
  margin-top: 16px;
}

.column-options {
  max-height: 320px;
  overflow-y: auto;
}

.column-checkbox-group {
  display: flex;
  flex-direction: column;
  gap: 6px;
  margin-top: 8px;
}

.column-checkbox {
  margin-right: 0;
}

.charts-section {
  margin-top: 20px;
}

.chart-container {
  width: 100%;
  margin: 20px 0;
}

.chart-container h4 {
  margin-bottom: 15px;
  color: #409eff;
  text-align: center;
}

.chart-wrapper {
  width: 100%;
  height: 400px;
  border: 1px solid #dcdfe6;
  border-radius: 4px;
}

@media (max-width: 768px) {
  .upload-row {
    flex-direction: column;
  }

  .date-section,
  .section-header,
  .section-actions {
    align-items: stretch;
    flex-direction: column;
  }
}
</style>
