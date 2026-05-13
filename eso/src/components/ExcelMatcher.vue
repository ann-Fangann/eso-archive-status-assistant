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
            <el-descriptions-item label="Delete行已回填数量">{{ summary.delete_completed_count }}</el-descriptions-item>
            <el-descriptions-item label="本次回填数量">{{ summary.matched_completed_count }}</el-descriptions-item>
            <el-descriptions-item label="截至日期已完成数量">{{ summary.due_completed_count }}</el-descriptions-item>
            <el-descriptions-item label="统计口径" :span="2">{{ summary.formula }}</el-descriptions-item>
          </el-descriptions>
        </div>

        <div v-if="unfinishedList.length" style="margin-top: 20px;">
          <div class="section-header">
            <h3>未完成清单</h3>
            <span>共 {{ unfinishedList.length }} 条</span>
          </div>
          <el-table :data="unfinishedList" style="width: 100%" stripe max-height="420">
            <el-table-column prop="零件号" label="零件号" min-width="130" />
            <el-table-column prop="操作类型" label="操作类型" width="90" />
            <el-table-column prop="功能组" label="功能组" width="120" />
            <el-table-column prop="工程师" label="工程师" min-width="160" />
            <el-table-column prop="ESO_Plan_Date" label="ESO Plan Date" min-width="130" />
            <el-table-column prop="ESO_Actual_Date" label="ESO Actual Date" min-width="140" />
            <el-table-column prop="ESO备注_类型" label="ESO备注" min-width="130" />
            <el-table-column prop="是否需要ESO_物流提供_" label="是否需要ESO" min-width="120" />
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
          <div v-if="uploadResult.sql_results.select_result && uploadResult.sql_results.select_result.data && !chartDataReady" style="margin-top: 20px; color: #f56c6c;">
            <p>图表数据处理失败，请检查控制台日志或联系管理员。</p>
          </div>
        </div>
      </div>
    </el-card>
  </div>
</template>

<script setup lang="ts">
import { ref, computed, nextTick, onUnmounted } from 'vue'
import axios from 'axios'
import { UploadFilled } from '@element-plus/icons-vue'
import type { UploadFile } from 'element-plus'
import * as echarts from 'echarts'

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

const getYesterday = () => {
  const day = new Date()
  day.setDate(day.getDate() - 1)
  const year = day.getFullYear()
  const month = String(day.getMonth() + 1).padStart(2, '0')
  const date = String(day.getDate()).padStart(2, '0')
  return `${year}-${month}-${date}`
}

const targetDate = ref(getYesterday())

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
      message.value = '无法连接到服务器，请确保后端服务已启动并在 http://localhost:8020 上运行'
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
const prepareChartData = (data: any[], selectResult: any = null) => {
  try {
    // 如果数据为空，直接返回
    if (!data || data.length === 0) {
      console.log('数据为空，无法生成图表')
      chartDataReady.value = false
      return
    }
    
    // 分析列信息
    columnInfo.value = analyzeColumnInfo(data)
    
    if (!columnInfo.value) {
      console.log('无法分析数据列信息')
      chartDataReady.value = false
      return
    }
    
    // 尝试找到功能组列名
    const functionGroupCol = columnInfo.value.availableColumns.find((col: string) => 
      col.toLowerCase().includes('功能组') || col.toLowerCase().includes('function') || col.toLowerCase().includes('group')
    ) 
    console.log('功能组列名检测结果:', functionGroupCol)
    
    // 尝试找到ESO_Actual_Date列名
    const esoActualDateCol = columnInfo.value.availableColumns.find((col: string) => 
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
    const nonEmptyCount = Number(selectResult?.summary?.completed_count ?? (data.length - fallbackEmptyCount))

    console.log('按功能组统计结果:', groupByFunction)
    console.log('空值数量:', emptyCount, '非空值数量:', nonEmptyCount)
    
    // 检查数据是否有效
    if (Object.keys(groupByFunction).length === 0 && emptyCount === 0 && nonEmptyCount === 0) {
      console.log('没有有效的数据用于生成图表')
      chartDataReady.value = false
      return
    }

    // 使用 nextTick 确保DOM元素渲染后再生成图表
    nextTick(() => {
      // 生成柱状图
      generateBarChart(groupByFunction)
      
      // 生成饼图
      generatePieChart(emptyCount, nonEmptyCount)
      
      chartDataReady.value = true
    })
  } catch (error) {
    console.error('准备图表数据时出错:', error)
    message.value = '生成图表数据时出错: ' + (error as Error).message
    messageType.value = 'error'
    chartDataReady.value = false
  }
}

// 生成柱状图
const generateBarChart = (data: Record<string, number>) => {
  // 等待下一个tick以确保DOM元素已渲染
  nextTick(() => {
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
  })
}

// 生成饼图
const generatePieChart = (emptyCount: number, nonEmptyCount: number) => {
  // 等待下一个tick以确保DOM元素已渲染
  nextTick(() => {
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
