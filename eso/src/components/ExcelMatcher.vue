<template>
  <div class="excel-matcher-container">
    <el-card shadow="hover" class="excel-card">
      <template #header>
        <div class="card-header">
          <span>Excel数据处理与可视化中心</span>
        </div>
      </template>
      
      <div class="upload-section">
        <div class="upload-row">
          <div class="upload-column">
            <h4>零件俱乐部-更新</h4>
            <el-upload
              drag
              :auto-upload="false"
              :on-change="(file: UploadFile) => handleChange(file, 'file1')"
              accept=".xlsx,.xls"
              style="margin-bottom: 20px;"
            >
              <el-icon class="el-icon--upload"><upload-filled /></el-icon>
              <div class="el-upload__text">
                将零件俱乐部-更新文件拖到此处，或 <em>点击上传</em>
              </div>
              <template #tip>
                <div class="el-upload__tip">
                  仅支持 .xlsx 或 .xls 格式的 Excel 文件
                </div>
              </template>
            </el-upload>
          </div>
          
          <div class="upload-column">
            <h4>ESO零件清单-更新</h4>
            <el-upload
              drag
              :auto-upload="false"
              :on-change="(file: UploadFile) => handleChange(file, 'file2')"
              accept=".xlsx,.xls"
              style="margin-bottom: 20px;"
            >
              <el-icon class="el-icon--upload"><upload-filled /></el-icon>
              <div class="el-upload__text">
                将ESO零件清单-更新文件拖到此处，或 <em>点击上传</em>
              </div>
              <template #tip>
                <div class="el-upload__tip">
                  仅支持 .xlsx 或 .xls 格式的 Excel 文件
                </div>
              </template>
            </el-upload>
          </div>
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
          :disabled="!uploadResult"
          :loading="refreshingUnfinished"
          @click="refreshUnfinishedList"
        >
          {{ refreshingUnfinished ? '正在生成...' : '重新生成未完成清单' }}
        </el-button>
      </div>
      
      <el-button
        type="success"
        @click="handleUpload"
        :disabled="!file1 || !file2"
        size="large"
        style="width: 100%;"
        :loading="uploading"
      >
        {{ uploading ? '正在上传和处理...' : '上传两个文件并执行数据处理' }}
      </el-button>
      
      <!-- 显示消息 -->
      <el-alert
        v-if="message"
        :title="message"
        :type="messageType"
        show-icon
        :closable="false"
        style="margin-top: 20px;"
      />
      
      <!-- 显示详细结果 -->
      <div v-if="uploadResult" style="margin-top: 20px;">
        <el-divider />
        <h3>导入结果</h3>
        <el-descriptions :column="1" border>
          <el-descriptions-item label="第一个文件名">{{ uploadResult.file1 }}</el-descriptions-item>
          <el-descriptions-item label="第二个文件名">{{ uploadResult.file2 }}</el-descriptions-item>
          <el-descriptions-item label="工作表数量">{{ uploadResult.sheets }}</el-descriptions-item>
        </el-descriptions>

        <el-alert
          v-if="uploadResult.persistence"
          :title="uploadResult.persistence.message"
          :type="uploadResult.persistence.saved ? 'success' : 'warning'"
          show-icon
          :closable="false"
          style="margin-top: 16px;"
        />

        <div v-if="uploadResult.modified_workbook" style="margin-top: 16px;">
          <el-alert
            v-if="uploadResult.modified_workbook.error"
            :title="uploadResult.modified_workbook.error"
            type="warning"
            show-icon
            :closable="false"
          />
          <el-button
            v-else
            type="primary"
            @click="downloadFile(uploadResult.modified_workbook.download_url)"
          >
            下载已回填的零件俱乐部
          </el-button>
        </div>

        <div v-if="summary" style="margin-top: 20px;">
          <h3>汇总结果</h3>
          <el-descriptions :column="2" border>
            <el-descriptions-item label="统计日期">{{ summary.target_date }}</el-descriptions-item>
            <el-descriptions-item label="有计划日期数量">{{ summary.planned_count }}</el-descriptions-item>
            <el-descriptions-item label="截至统计日期计划数量">{{ summary.due_planned_count }}</el-descriptions-item>
            <el-descriptions-item label="有实际完成日期数量">{{ summary.completed_count }}</el-descriptions-item>
            <el-descriptions-item label="未完成数量">{{ summary.unfinished_count }}</el-descriptions-item>
            <el-descriptions-item label="Delete行已有实际日期数量">{{ summary.delete_completed_count }}</el-descriptions-item>
            <el-descriptions-item label="Delete行排除未完成数量">{{ summary.delete_unfinished_excluded_count || 0 }}</el-descriptions-item>
            <el-descriptions-item label="本次回填数量">{{ summary.matched_completed_count }}</el-descriptions-item>
            <el-descriptions-item label="截至日期已完成数量">{{ summary.due_completed_count }}</el-descriptions-item>
            <el-descriptions-item label="统计口径" :span="2">{{ summary.formula }}</el-descriptions-item>
          </el-descriptions>

          <el-alert
            v-if="duplicateWarningText"
            :title="duplicateWarningText"
            type="warning"
            show-icon
            :closable="false"
            style="margin-top: 12px;"
          />

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
        </div>

        <div v-if="unfinishedList.length" style="margin-top: 20px;">
          <div class="section-header">
            <h3>未完成清单</h3>
            <div class="section-actions">
              <span>共 {{ unfinishedList.length }} 条</span>
              <el-popover
                placement="bottom"
                title="选择显示和导出的列"
                width="320"
                trigger="click"
              >
                <div class="column-options">
                  <el-checkbox
                    :model-value="isAllUnfinishedColumnsSelected"
                    :indeterminate="isSomeUnfinishedColumnsSelected"
                    @change="toggleUnfinishedColumns"
                  >
                    全选/全不选
                  </el-checkbox>
                  <el-checkbox-group
                    v-model="selectedUnfinishedColumns"
                    class="column-checkbox-group"
                    @change="saveUnfinishedColumnSelection"
                  >
                    <el-checkbox
                      v-for="col in allUnfinishedColumns"
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
                :disabled="!selectedUnfinishedColumns.length"
                @click="exportUnfinishedList"
              >
                导出未完成清单
              </el-button>
            </div>
          </div>
          <el-table :data="unfinishedDisplayRows" style="width: 100%" stripe max-height="420">
            <el-table-column
              v-for="col in selectedUnfinishedColumnDefs"
              :key="col.prop"
              :prop="col.prop"
              :label="col.label"
              :min-width="col.width"
              show-overflow-tooltip
            />
          </el-table>
        </div>
        
        <h4 style="margin-top: 20px;">各工作表详情</h4>
        <el-table :data="resultTableData" style="width: 100%" stripe>
          <el-table-column prop="sheet" label="工作表名" />
          <el-table-column prop="result" label="导入结果" />
        </el-table>
        
        <!-- SQL执行结果显示 -->
        <div v-if="uploadResult.sql_results" style="margin-top: 20px;">
          <!-- <el-divider />
          <h3>SQL执行结果</h3>
          <div style="margin-bottom: 20px;">
            <h4>更新操作结果</h4>
            <p>{{ uploadResult.sql_results.update_result }}</p>
          </div> -->
          
          <!-- 图表展示区域 -->
          <div v-if="chartDataReady" style="margin-top: 20px;">
            <el-divider />
            <h3>数据可视化分析</h3>
            
            <!-- 显示ESO_Actual_Date为空的总记录数 -->
            <!-- <div style="margin-bottom: 20px;">
              <el-alert
                :title="`ESO_Actual_Date为空的总记录数: ${uploadResult.sql_results.select_result.total_empty_count || 0}`"
                type="info"
                show-icon
                :closable="false"
              />
            </div> -->
            
            <!-- 柱状图 -->
            <div class="chart-container">
              <h4>功能组ESO计划未完成分布</h4>
              <div ref="barChartRef" class="chart-wrapper"></div>
            </div>
            
            <!-- 饼图 -->
            <div class="chart-container" style="margin-top: 30px;">
              <h4>ESO执行状态数据分布</h4>
              <div ref="pieChartRef" class="chart-wrapper"></div>
            </div>
          </div>
          
          <!-- 如果图表数据未准备好，显示提示信息 -->
          <div v-if="uploadResult.sql_results.select_result && uploadResult.sql_results.select_result.error" style="margin-top: 20px; color: #f56c6c;">
            <p>{{ uploadResult.sql_results.select_result.error }}</p>
          </div>
        </div>
      </div>
    </el-card>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, nextTick, onUnmounted, watch } from 'vue'
import axios from 'axios'
import { UploadFilled } from '@element-plus/icons-vue'
import { ElMessage } from 'element-plus'
import type { UploadFile } from 'element-plus'
import * as echarts from 'echarts'
import * as XLSX from 'xlsx-js-style'

const file1 = ref<File | null>(null)
const file2 = ref<File | null>(null)
const message = ref('')
const messageType = ref<'success' | 'error' | 'warning' | 'info'>('info')
const uploadResult = ref<any>(null)
const uploading = ref(false)
const refreshingUnfinished = ref(false)

// 图表相关引用
const barChartRef = ref<HTMLElement | null>(null)
const pieChartRef = ref<HTMLElement | null>(null)
const chartDataReady = ref(false)

// 添加columnInfo响应式变量
const columnInfo = ref<any>(null)

// 计算属性，将结果转换为表格数据
const resultTableData = computed(() => {
  if (!uploadResult.value || !uploadResult.value.results) return []
  
  return Object.entries(uploadResult.value.results).map(([sheet, result]) => ({
    sheet,
    result
  }))
})

const summary = computed(() => uploadResult.value?.sql_results?.select_result?.summary || null)
const unfinishedList = computed(() => uploadResult.value?.sql_results?.select_result?.data || [])
const groupStats = computed(() => uploadResult.value?.sql_results?.select_result?.group_stats || [])
const duplicateStats = computed(() => uploadResult.value?.sql_results?.select_result?.duplicate_stats || null)

const duplicateWarningText = computed(() => {
  const sheetDuplicateCount = duplicateStats.value?.sheet?.duplicate_part_count || 0
  const archiveDuplicateCount = duplicateStats.value?.sheet1?.duplicate_part_count || 0
  const warnings = []
  if (sheetDuplicateCount) warnings.push(`零件俱乐部存在 ${sheetDuplicateCount} 个重复零件号`)
  if (archiveDuplicateCount) warnings.push(`ESO零件清单存在 ${archiveDuplicateCount} 个重复零件号，回填时按同一零件号的最大归档日期处理`)
  return warnings.join('；')
})

const reportText = computed(() => {
  if (!summary.value) return ''

  const topGroups = groupStats.value
    .slice(0, 5)
    .map((item: any) => `${item['功能组'] || '未知功能组'} ${item.count || 0} 项`)
    .join('，')
  const excludedDeleteCount = summary.value.delete_unfinished_excluded_count || 0
  const matchedCount = summary.value.matched_completed_count ?? 0

  return [
    `截至 ${summary.value.target_date}，ESO应完成计划 ${summary.value.due_planned_count || 0} 项，已完成 ${summary.value.due_completed_count || 0} 项，未完成 ${summary.value.unfinished_count || 0} 项。`,
    `本次根据ESO零件清单回填实际归档日期 ${matchedCount} 项；未完成统计已排除操作类型为 D 的记录${excludedDeleteCount ? `，本次排除 ${excludedDeleteCount} 项` : ''}。`,
    topGroups ? `未完成主要分布：${topGroups}。` : '当前统计日期没有未完成项。',
  ].join('\n')
})

const columnLabelMap: Record<string, string> = {
  'ESO备注_类型': 'ESO备注',
  '首次申请项目': '首次应用项目',
  '首次应用项目': '首次应用项目',
  'ESO_Plan_Date': 'ESO计划',
  'ESO_Actual_Date': 'ESO实际归档日期',
  '是否需要ESO_物流提供_': '是否需要ESO',
  't_2D_Drawing_No_': '2D Drawing No.',
  'ESO状态': 'ESO状态'
}

const defaultUnfinishedColumns = [
  '零件号',
  'FFC中文描述',
  '工程师',
  '功能组',
  '首次申请项目',
  '首次应用项目',
  'ESO_Plan_Date',
  'ESO_Actual_Date'
]

const selectedUnfinishedColumns = ref<string[]>([])

const withEsoStatus = (row: Record<string, any>) => ({
  ...row,
  ESO状态: '审批中'
})

const unfinishedDisplayRows = computed(() => unfinishedList.value.map((row: Record<string, any>) => withEsoStatus(row)))

const getColumnLabel = (prop: string) => columnLabelMap[prop] || prop

const getColumnWidth = (prop: string) => {
  if (['零件号', '工程师', 'FFC中文描述', 'ESO备注_类型', '备注', '首次申请项目'].includes(prop)) return 160
  if (prop.includes('Date') || prop.includes('日期')) return 140
  if (['操作类型', '车型年', 'ESO状态'].includes(prop)) return 100
  return 120
}

const allUnfinishedColumns = computed(() => {
  const keySet = new Set<string>()
  unfinishedDisplayRows.value.forEach((row: Record<string, any>) => {
    Object.keys(row).forEach(key => keySet.add(key))
  })

  const defaultKeys = defaultUnfinishedColumns.filter(key => keySet.has(key))
  const esoKeys = Array.from(keySet).filter(key => {
    const isDefault = defaultKeys.includes(key)
    const isEsoColumn = key.toUpperCase().includes('ESO')
    return !isDefault && isEsoColumn
  })

  const orderedKeys = [
    ...defaultKeys,
    ...esoKeys.sort((a, b) => a.localeCompare(b, 'zh-Hans-CN'))
  ]

  return orderedKeys.map(prop => ({
    prop,
    label: getColumnLabel(prop),
    width: getColumnWidth(prop)
  }))
})

const selectedUnfinishedColumnDefs = computed(() => {
  const selectedSet = new Set(selectedUnfinishedColumns.value)
  return allUnfinishedColumns.value.filter(col => selectedSet.has(col.prop))
})

const isAllUnfinishedColumnsSelected = computed(() => {
  return allUnfinishedColumns.value.length > 0 && selectedUnfinishedColumns.value.length === allUnfinishedColumns.value.length
})

const isSomeUnfinishedColumnsSelected = computed(() => {
  return selectedUnfinishedColumns.value.length > 0 && selectedUnfinishedColumns.value.length < allUnfinishedColumns.value.length
})

const getYesterday = () => {
  const day = new Date()
  day.setDate(day.getDate() - 1)
  const year = day.getFullYear()
  const month = String(day.getMonth() + 1).padStart(2, '0')
  const date = String(day.getDate()).padStart(2, '0')
  return `${year}-${month}-${date}`
}

const targetDate = ref(getYesterday())

const getSavedUnfinishedColumnSelection = () => {
  try {
    const saved = localStorage.getItem('eso_unfinished_columns_v2')
    return saved ? JSON.parse(saved) : []
  } catch (error) {
    console.error('读取未完成清单列选择失败:', error)
    return []
  }
}

const saveUnfinishedColumnSelection = () => {
  localStorage.setItem('eso_unfinished_columns_v2', JSON.stringify(selectedUnfinishedColumns.value))
}

const resetUnfinishedColumnSelection = () => {
  const allKeys = allUnfinishedColumns.value.map(col => col.prop)
  if (!allKeys.length) {
    selectedUnfinishedColumns.value = []
    return
  }

  const saved = getSavedUnfinishedColumnSelection().filter((key: string) => allKeys.includes(key))
  const defaults = defaultUnfinishedColumns.filter(key => allKeys.includes(key))
  selectedUnfinishedColumns.value = saved.length ? saved : defaults
}

const toggleUnfinishedColumns = (checked: boolean) => {
  selectedUnfinishedColumns.value = checked ? allUnfinishedColumns.value.map(col => col.prop) : []
  saveUnfinishedColumnSelection()
}

watch(allUnfinishedColumns, (columns) => {
  const allKeys = columns.map(col => col.prop)
  const validSelected = selectedUnfinishedColumns.value.filter(key => allKeys.includes(key))

  if (!allKeys.length) {
    selectedUnfinishedColumns.value = []
    return
  }

  if (!validSelected.length) {
    resetUnfinishedColumnSelection()
    return
  }

  if (validSelected.length !== selectedUnfinishedColumns.value.length) {
    selectedUnfinishedColumns.value = validSelected
    saveUnfinishedColumnSelection()
  }
}, { immediate: true })

const handleChange = (uploadFile: UploadFile, fileKey: string) => {
  if (fileKey === 'file1') {
    file1.value = uploadFile.raw || null
  } else if (fileKey === 'file2') {
    file2.value = uploadFile.raw || null
  }
  
  // 重置之前的结果
  message.value = ''
  uploadResult.value = null
  chartDataReady.value = false
}

const handleUpload = async () => {
  if (!file1.value || !file2.value) {
    message.value = '请上传两个Excel文件'
    messageType.value = 'warning'
    return
  }

  const formData = new FormData()
  formData.append('file1', file1.value)
  formData.append('file2', file2.value)
  if (targetDate.value) {
    formData.append('target_date', targetDate.value)
  }

  try {
    uploading.value = true
    message.value = '正在上传和处理文件...'
    messageType.value = 'info'
    
    console.log('开始上传两个文件到 /api/upload-two-excel/')
    
    // 使用代理地址而不是直接地址
    const response = await axios.post('/api/upload-two-excel/', formData, {
      headers: {
        'Content-Type': 'multipart/form-data'
      },
      timeout: 600000 // 10分钟超时，处理两个文件可能需要更长时间
    })
    
    console.log('收到服务器响应:', response.data)
    message.value = '文件上传并处理成功！'
    messageType.value = 'success'
    uploadResult.value = response.data
    
    // 如果有查询结果，准备图表数据
    if (response.data.sql_results && response.data.sql_results.select_result && response.data.sql_results.select_result.data) {
      await nextTick()
      await prepareChartData(response.data.sql_results.select_result.data, response.data.sql_results.select_result)
    }
    
  } catch (error: any) {
    console.error('上传失败:', error)
    
    // 根据错误类型显示不同的消息
    if (error.code === 'ECONNABORTED') {
      message.value = '请求超时，请稍后再试或联系管理员'
    } else if (!error.response) {
      message.value = '无法连接到服务器，请确保后端服务已启动并在 http://localhost:8030 上运行'
    } else {
      message.value = '上传失败: ' + (error.response?.data?.detail || error.message || '未知错误')
    }
    
    messageType.value = 'error'
    uploadResult.value = null
    chartDataReady.value = false
  } finally {
    uploading.value = false
  }
}

const refreshUnfinishedList = async () => {
  if (!uploadResult.value) {
    message.value = '请先上传并处理两个Excel文件'
    messageType.value = 'warning'
    return
  }

  try {
    refreshingUnfinished.value = true
    message.value = '正在按统计日期生成未完成清单...'
    messageType.value = 'info'
    chartDataReady.value = false

    const response = await axios.post('/api/unfinished-list/', {
      target_date: targetDate.value
    }, {
      timeout: 120000
    })

    const previousMatchedCount = summary.value?.matched_completed_count
    if (!uploadResult.value.sql_results) {
      uploadResult.value.sql_results = {}
    }
    uploadResult.value.sql_results.select_result = response.data.select_result
    if (response.data.persistence) {
      uploadResult.value.persistence = response.data.persistence
    }

    const currentSummary = uploadResult.value.sql_results.select_result?.summary
    if (currentSummary && currentSummary.matched_completed_count == null && previousMatchedCount != null) {
      currentSummary.matched_completed_count = previousMatchedCount
    }

    message.value = '未完成清单已重新生成'
    messageType.value = 'success'

    if (response.data.select_result?.data) {
      await nextTick()
      await prepareChartData(response.data.select_result.data, response.data.select_result)
    }
  } catch (error: any) {
    console.error('重新生成未完成清单失败:', error)
    message.value = '重新生成未完成清单失败: ' + (error.response?.data?.detail || error.message || '未知错误')
    messageType.value = 'error'
  } finally {
    refreshingUnfinished.value = false
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

const getExportDateString = () => {
  const dateValue = summary.value?.target_date || targetDate.value || getYesterday()
  return String(dateValue).replace(/-/g, '')
}

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
        alignment: {
          vertical: 'center',
          wrapText: true
        }
      }

      if (isHeader) {
        cell.s.fill = { fgColor: { rgb: headerColor } }
        cell.s.font = { color: { rgb: 'FFFFFF' }, bold: true }
      }
    }
  }
}

const appendJsonSheet = (
  wb: XLSX.WorkBook,
  rows: Record<string, any>[],
  sheetName: string,
  widths: number[] = []
) => {
  const ws = XLSX.utils.json_to_sheet(rows.length ? rows : [{}])
  styleWorksheet(ws)
  if (widths.length) {
    ws['!cols'] = widths.map(wch => ({ wch }))
  }
  XLSX.utils.book_append_sheet(wb, ws, sheetName)
}

const exportUnfinishedList = () => {
  if (!unfinishedDisplayRows.value.length) {
    ElMessage.warning('当前没有可导出的未完成清单')
    return
  }

  if (!selectedUnfinishedColumns.value.length) {
    ElMessage.warning('请至少选择一列')
    return
  }

  const selectedDefs = selectedUnfinishedColumnDefs.value
  const exportRows = unfinishedDisplayRows.value.map((row: Record<string, any>) => {
    const output: Record<string, any> = {}
    selectedDefs.forEach(col => {
      output[col.label] = row[col.prop] ?? ''
    })
    return output
  })

  const wb = XLSX.utils.book_new()
  const ws = XLSX.utils.json_to_sheet(exportRows)
  styleWorksheet(ws)
  ws['!cols'] = selectedDefs.map(col => ({ wch: Math.min(Math.max(Math.round((col.width || 120) / 8), 12), 28) }))
  XLSX.utils.book_append_sheet(wb, ws, '延期未发布ESO零件')

  const summaryRows = summary.value ? [
    { 指标: '统计日期', 数值: summary.value.target_date || '' },
    { 指标: '有计划日期数量', 数值: summary.value.planned_count || 0 },
    { 指标: '截至统计日期计划数量', 数值: summary.value.due_planned_count || 0 },
    { 指标: '有实际完成日期数量', 数值: summary.value.completed_count || 0 },
    { 指标: '截至日期已完成数量', 数值: summary.value.due_completed_count || 0 },
    { 指标: '未完成数量', 数值: summary.value.unfinished_count || 0 },
    { 指标: '本次回填数量', 数值: summary.value.matched_completed_count ?? 0 },
    { 指标: 'Delete行已有实际日期数量', 数值: summary.value.delete_completed_count || 0 },
    { 指标: 'Delete行排除未完成数量', 数值: summary.value.delete_unfinished_excluded_count || 0 },
    { 指标: '统计口径', 数值: summary.value.formula || '' },
  ] : []
  appendJsonSheet(wb, summaryRows, '汇总结果', [24, 80])

  const statField = summary.value?.group_field || '功能组'
  const statRows = groupStats.value.map((item: any) => ({
    统计维度: statField,
    名称: item['功能组'] || '未知功能组',
    未完成数量: item.count || 0
  }))
  appendJsonSheet(wb, statRows, '部门统计', [16, 32, 14])

  const duplicateRows = [
    ...(duplicateStats.value?.sheet?.samples || []).map((item: any) => ({
      来源: '零件俱乐部',
      零件号: item['零件号'],
      重复次数: item.count
    })),
    ...(duplicateStats.value?.sheet1?.samples || []).map((item: any) => ({
      来源: 'ESO零件清单',
      零件号: item['零件号'],
      重复次数: item.count
    })),
  ]
  if (duplicateRows.length) {
    appendJsonSheet(wb, duplicateRows, '重复零件号', [18, 28, 12])
  }

  XLSX.writeFile(wb, `延期未发布ESO零件-${getExportDateString()}.xlsx`)
  ElMessage.success('未完成清单已导出')
}

// 分析数据列信息
const analyzeColumnInfo = (data: any[]) => {
  if (data.length === 0) return null
  
  const firstRow = data[0]
  const availableColumns = Object.keys(firstRow)
  
  const hasFunctionGroup = availableColumns.some(col => 
    col.toLowerCase().includes('功能组') || col.toLowerCase().includes('function') || col.toLowerCase().includes('group')
  )
  
  const hasEsoActualDate = availableColumns.some(col => 
    col.toLowerCase().includes('actual') || col.toLowerCase().includes('date') || col.includes('ESO_Actual_Date')
  )
  
  return {
    availableColumns,
    hasFunctionGroup,
    hasEsoActualDate
  }
}

// 准备图表数据
const prepareChartData = async (data: any[] = [], selectResult: any = null) => {
  try {
    data = data || []
    
    // 分析列信息
    columnInfo.value = analyzeColumnInfo(data)

    const availableColumns = columnInfo.value?.availableColumns || []
    
    // 尝试找到功能组列名
    const functionGroupCol = availableColumns.find((col: string) => 
      col.toLowerCase().includes('功能组') || col.toLowerCase().includes('function') || col.toLowerCase().includes('group')
    ) 
    console.log('功能组列名检测结果:', functionGroupCol)
    
    // 尝试找到ESO_Actual_Date列名
    const esoActualDateCol = availableColumns.find((col: string) => 
      col.toLowerCase().includes('actual') || col.toLowerCase().includes('date') || col.includes('ESO_Actual_Date')
    ) 
    console.log('ESO_Actual_Date列名检测结果:', esoActualDateCol)
    
    // 如果找不到对应列，使用所有可能的列名进行尝试
    if (!functionGroupCol) {
      console.warn('未找到明确的功能组列，将尝试常见的功能组列名')
    }
    
    if (!esoActualDateCol) {
      console.warn('未找到明确的ESO_Actual_Date列，将尝试常见的日期列名')
    }
    
    // 按功能组分组统计未完成数量。优先使用后端按业务口径统计后的结果。
    const groupByFunction: Record<string, number> = {}

    if (selectResult?.group_stats?.length) {
      selectResult.group_stats.forEach((item: any) => {
        const groupName = item['功能组'] || item['部门'] || '未知功能组'
        groupByFunction[String(groupName)] = Number(item.count || item['延期未发布数量'] || 0)
      })
    }
    
    if (Object.keys(groupByFunction).length === 0) data.forEach((item, index) => {
      // 根据检测到的列名或常见列名获取功能组值
      let funcGroup = null
      if (functionGroupCol) {
        funcGroup = item[functionGroupCol]
      } else {
        // 尝试常见的功能组列名
        funcGroup = item['功能组'] || item['function_group'] || item['functionGroup'] || item['Function Group'] || '未知功能组'
      }
      
      // 如果funcGroup是null或undefined，设置为'未知功能组'
      if (funcGroup === null || funcGroup === undefined || funcGroup === '') {
        funcGroup = '未知功能组'
      } else {
        // 确保funcGroup是字符串类型
        funcGroup = String(funcGroup)
      }
      
      // 根据检测到的列名或常见列名检查ESO_Actual_Date是否为空
      let esoActualDateValue = null
      if (esoActualDateCol) {
        esoActualDateValue = item[esoActualDateCol]
      } else {
        // 尝试常见的日期列名
        esoActualDateValue = item['ESO_Actual_Date'] || item['eso_actual_date'] || item['esoActualDate'] || 
                             item['Actual_Date'] || item['actual_date'] || item['ActualDate'] || 
                             item['Date'] || item['date']
      }
      
      // 检查ESO_Actual_Date是否为空
      if (!esoActualDateValue || esoActualDateValue === '' || esoActualDateValue === null || esoActualDateValue === undefined) {
        if (!groupByFunction[funcGroup]) {
          groupByFunction[funcGroup] = 0
        }
        groupByFunction[funcGroup]++
      }
    })

    // 统计ESO_Actual_Date为空和非空的数量
    const fallbackEmptyCount = data.filter(item => {
      let esoActualDateValue = null
      if (esoActualDateCol) {
        esoActualDateValue = item[esoActualDateCol]
      } else {
        esoActualDateValue = item['ESO_Actual_Date'] || item['eso_actual_date'] || item['esoActualDate'] || 
                             item['Actual_Date'] || item['actual_date'] || item['ActualDate'] || 
                             item['Date'] || item['date']
      }
      return !esoActualDateValue || esoActualDateValue === '' || esoActualDateValue === null || esoActualDateValue === undefined
    }).length
    
    const emptyCount = Number(selectResult?.summary?.unfinished_count ?? fallbackEmptyCount)
    const nonEmptyCount = Number(selectResult?.summary?.due_completed_count ?? selectResult?.summary?.completed_count ?? (data.length - fallbackEmptyCount))

    console.log('按功能组统计结果:', groupByFunction)
    console.log('空值数量:', emptyCount, '非空值数量:', nonEmptyCount)

    chartDataReady.value = true
    await nextTick()
    generateBarChart(groupByFunction)
    generatePieChart(emptyCount, nonEmptyCount)
  } catch (error) {
    console.error('准备图表数据时出错:', error)
    message.value = '生成图表数据时出错: ' + (error as Error).message
    messageType.value = 'error'
    chartDataReady.value = false
  }
}

// 生成柱状图
const generateBarChart = (data: Record<string, number>) => {
  if (!barChartRef.value) {
    console.error('柱状图容器不存在', barChartRef.value)
    return
  }
    
  // 销毁已存在的实例
  const existingChart = echarts.getInstanceByDom(barChartRef.value)
  if (existingChart) {
    existingChart.dispose()
  }
    
  // 初始化图表
  const chart = echarts.init(barChartRef.value)
    
  // 如果没有数据，显示提示
  if (Object.keys(data).length === 0) {
    chart.setOption({
      title: {
        text: '该统计日期没有未完成项',
        left: 'center',
        top: 'center',
        textStyle: {
          fontSize: 16
        }
      }
    })
    return
  }
    
  const option = {
    title: {
      text: '各功能组未完成的数量',
      left: 'center'
    },
    tooltip: {
      trigger: 'axis',
      axisPointer: {
        type: 'shadow'
      }
    },
    xAxis: {
      type: 'category',
      data: Object.keys(data),
      axisLabel: {
        interval: 0,
        rotate: 45,
        fontSize: 12
      }
    },
    yAxis: {
      type: 'value',
      name: '数量'
    },
    series: [{
      data: Object.values(data),
      type: 'bar',
      itemStyle: {
        color: '#409eff'
      }
    }]
  }
    
  chart.setOption(option)
    
  // 响应式调整
  window.addEventListener('resize', () => {
    if (barChartRef.value) {
      const chart = echarts.getInstanceByDom(barChartRef.value)
      if (chart) chart.resize()
    }
  })
}

// 生成饼图
const generatePieChart = (emptyCount: number, nonEmptyCount: number) => {
  if (!pieChartRef.value) {
    console.error('饼图容器不存在', pieChartRef.value)
    return
  }
    
  // 销毁已存在的实例
  const existingChart = echarts.getInstanceByDom(pieChartRef.value)
  if (existingChart) {
    existingChart.dispose()
  }
    
  // 初始化图表
  const chart = echarts.init(pieChartRef.value)
    
    // 如果没有数据，显示提示
    if (emptyCount === 0 && nonEmptyCount === 0) {
      chart.setOption({
        title: {
          text: '没有可显示的数据',
          left: 'center',
          top: 'center',
          textStyle: {
            fontSize: 16
          }
        }
      })
    return
  }
    
    const option = {
      title: {
        text: '未完成数量占比',
        left: 'center'
      },
      tooltip: {
        trigger: 'item',
        formatter: '{a} <br/>{b}: {c} ({d}%)'
      },
      legend: {
        orient: 'horizontal',
        bottom: 'bottom'
      },
      series: [
        {
          name: '状态',
          type: 'pie',
          radius: ['40%', '70%'],
          avoidLabelOverlap: false,
          itemStyle: {
            borderRadius: 10,
            borderColor: '#fff',
            borderWidth: 2
          },
          label: {
            show: true,
            formatter: '{b}: {c} ({d}%)'
          },
          emphasis: {
            label: {
              show: true,
              fontSize: 16,
              fontWeight: 'bold'
            }
          },
          data: [
            { value: emptyCount, name: '未完成', itemStyle: { color: '#f56c6c' } },
            { value: nonEmptyCount, name: '已完成', itemStyle: { color: '#67c23a' } }
          ]
        }
      ]
    }
    
  chart.setOption(option)
    
  // 响应式调整
  window.addEventListener('resize', () => {
    if (pieChartRef.value) {
      const chart = echarts.getInstanceByDom(pieChartRef.value)
      if (chart) chart.resize()
    }
  })
}

// 组件卸载时销毁图表实例
onUnmounted(() => {
  if (barChartRef.value) {
    const chart = echarts.getInstanceByDom(barChartRef.value)
    if (chart) chart.dispose()
  }
  if (pieChartRef.value) {
    const chart = echarts.getInstanceByDom(pieChartRef.value)
    if (chart) chart.dispose()
  }
})
</script>

<style scoped>
.excel-matcher-container {
  height: 100%;
  display: flex;
  flex-direction: column;
}

.excel-card {
  flex: 1;
  display: flex;
  flex-direction: column;
  height: 100%;
}

:deep(.el-card__body) {
  flex: 1;
  display: flex;
  flex-direction: column;
}

.card-header {
  font-size: 18px;
  font-weight: bold;
}

.upload-section {
  margin-bottom: 20px;
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
  margin-bottom: 12px;
}

.section-header h3 {
  margin: 0;
}

.section-header span {
  color: #606266;
  font-size: 14px;
}

.section-actions {
  display: flex;
  align-items: center;
  gap: 10px;
  flex-wrap: wrap;
  justify-content: flex-end;
}

.report-panel {
  margin-top: 16px;
}

.report-panel h4 {
  margin: 0;
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

.upload-row {
  display: flex;
  gap: 20px;
}

.upload-column {
  flex: 1;
  min-width: 0; /* 防止内容溢出 */
}

.upload-column h4 {
  margin-top: 0;
  margin-bottom: 15px;
  text-align: center;
  color: #409eff;
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
</style>
