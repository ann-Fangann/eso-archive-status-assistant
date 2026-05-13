<template>
  <div class="chat-container-wrapper">
    <div class="chat-container">
      <div class="chat-history" ref="chatHistoryRef">
        <div 
          v-for="(message, index) in messages" 
          :key="index" 
          :class="['message', message.type]"
        >
          <div class="message-content">
            <div class="sender">{{ message.sender }}</div>
            <!-- 检查是否是表格数据 -->
            <div v-if="message.tableData" class="table-container">
              <div class="table-summary">
                共查询到 {{ message.tableData.total }} 条数据
              </div>
              
              <!-- 列选择器 -->
              <div class="column-selector">
                <el-popover
                  placement="bottom"
                  title="选择要导出的列"
                  width="300"
                  trigger="click"
                >
                  <div class="column-options">
                    <el-checkbox
                      v-model="message.tableData.exportAllColumns"
                      @change="toggleAllColumns($event, message.tableData)"
                    >
                      全选/全不选
                    </el-checkbox>
                    <el-checkbox-group
                      v-model="message.tableData.selectedColumns"
                      class="column-checkbox-group"
                      @change="(val) => handleColumnChange(val, message.tableData)"
                    >
                      <el-checkbox
                        v-for="col in message.tableData.columns"
                        :key="col.prop"
                        :label="col.prop"
                        class="column-checkbox"
                      >
                        {{ col.label }}
                      </el-checkbox>
                    </el-checkbox-group>
                  </div>
                  <template #reference>
                    <el-button 
                      size="small" 
                      type="primary" 
                      icon="Setting"
                    >
                      选择列
                    </el-button>
                  </template>
                </el-popover>
                
                <!-- 导出按钮 -->
                <el-button 
                  size="small" 
                  type="primary" 
                  @click="exportToExcel(message.tableData.records, message.tableData.selectedColumns, message.tableData.exportAllColumns)"
                  icon="Download"
                >
                  导出为Excel
                </el-button>
              </div>
              
              <el-table 
                :data="message.tableData.displayRecords" 
                style="width: 100%" 
                size="small"
                stripe
                :height="message.tableData.height || 300"
                max-height="400"
              >
                <el-table-column
                  v-for="col in getOrderedColumns(message.tableData)"
                  :key="col.prop"
                  :prop="col.prop"
                  :label="col.label"
                  :min-width="col.width || 120"
                />
              </el-table>
              <div class="pagination-container" v-if="message.tableData.pagination">
                <el-pagination
                  layout="prev, pager, next"
                  :total="message.tableData.total"
                  :page-size="message.tableData.pageSize"
                  :current-page="message.tableData.currentPage"
                  @current-change="(page) => handlePageChange(index, page)"
                />
              </div>
            </div>
            <!-- 普通文本消息 -->
            <div v-else class="text">{{ message.text }}</div>
            
            <!-- 显示提问样例，仅在欢迎消息中显示 -->
            <div v-if="message.examples && message.examples.length" class="examples-container">
              <el-collapse accordion>
                <el-collapse-item title="点击展开提问样例" name="examples">
                  <div v-for="(example, idx) in message.examples" :key="idx" class="example-item">
                    <div class="example-question">
                      <strong>提问：</strong>{{ example.question }}
                    </div>
                    <div class="example-image" v-if="example.image">
                      <img :src="example.image" alt="回答样例" @click="showImagePreview(example.image)" class="clickable-image" @error="handleImageError" />
                    </div>
                    <div class="example-placeholder" v-else>
                      这里将展示图像样例
                    </div>
                  </div>
                </el-collapse-item>
              </el-collapse>
            </div>
            
            <div class="timestamp">{{ message.timestamp }}</div>
          </div>
        </div>
      </div>
      <div class="input-area">
        <el-input
          v-model="inputMessage"
          type="textarea"
          :rows="3"
          placeholder="请输入您的问题..."
          @keyup.enter.exact="sendMessage"
        />
        <el-button 
          type="primary" 
          @click="sendMessage" 
          :loading="sending"
          class="send-button"
        >
          发送
        </el-button>
      </div>
    </div>
  </div>
  
  <!-- 图像预览组件 -->
  <el-image-viewer 
    v-if="showImageViewer" 
    :url-list="[imageViewerUrl]" 
    @close="showImageViewer = false" 
  />
</template>

<script setup>
import { ref, nextTick } from 'vue'
import { ElMessage, ElImageViewer } from 'element-plus'
import { runWorkflow } from '../utils/dify.js'
import * as XLSX from 'xlsx-js-style'
import exampleImage from '../assets/image.png' // 导入示例图像

const messages = ref([])
const inputMessage = ref('')
const sending = ref(false)
const chatHistoryRef = ref(null)
const showImageViewer = ref(false)
const imageViewerUrl = ref('')

const scrollToBottom = () => {
  nextTick(() => {
    if (chatHistoryRef.value) {
      chatHistoryRef.value.scrollTop = chatHistoryRef.value.scrollHeight
    }
  })
}

// 添加欢迎消息，包含可折叠的提问案例
messages.value.push({
  type: 'received',
  sender: 'AI助手',
  text: `您好！我是您的智能助手，可以帮您查询数据库中的信息。`,
  examples: [
    {
      question: "请帮我查询在2024年4月5号之前的数据",
      image: exampleImage  // 使用导入的图像
    }
  ],
  timestamp: new Date().toLocaleTimeString(),
  // 确保消息对象包含时间戳
})

// 处理分页变化
const handlePageChange = (messageIndex, page) => {
  const message = messages.value[messageIndex]
  if (message && message.tableData) {
    // 保存现有的列选择状态
    const { selectedColumns, exportAllColumns } = message.tableData
    
    // 更新当前页码
    message.tableData.currentPage = page
    
    // 计算当前页应该显示的数据
    const start = (page - 1) * message.tableData.pageSize
    const end = start + message.tableData.pageSize
    message.tableData.displayRecords = message.tableData.records.slice(start, end)
    
    // 恢复列选择状态
    message.tableData.selectedColumns = selectedColumns
    message.tableData.exportAllColumns = exportAllColumns
  }
}

// 导出为Excel
const exportToExcel = (data, selectedColumns = null, exportAllColumns = true) => {
  try {
    // 创建工作簿
    const wb = XLSX.utils.book_new();

    // 如果有选择的列，则按照selectedColumns的顺序导出
    let exportData = data;
    if (selectedColumns && !exportAllColumns && selectedColumns.length > 0) {
      // 分离普通列和统计列
      const statsColumns = ['部门', '延期未发布数量'];
      const normalColumns = selectedColumns.filter(col => !statsColumns.includes(col));
      const statsColsSelected = selectedColumns.filter(col => statsColumns.includes(col));

      // 重新排列：普通列 + 空列1 + 空列2 + 统计列
      let reorderedColumns = [...normalColumns];
      if (statsColsSelected.length > 0) {
        reorderedColumns = [...normalColumns, '', '', ...statsColsSelected];
      }

      exportData = data.map(row => {
        const filteredRow = {};
        reorderedColumns.forEach((col, index) => {
          if (col === '') {
            // 空列使用空字符串作为key
            filteredRow[`__empty_${index}`] = '';
          } else {
            filteredRow[col] = row[col];
          }
        });
        return filteredRow;
      });
    }

    // 将数据转换为工作表
    const ws = XLSX.utils.json_to_sheet(exportData);

    // 清除空列的列名（设置为空字符串）
    const range = XLSX.utils.decode_range(ws['!ref']);

    // 遍历第一行（列名行），将空列的列名设置为空
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
      if (ws[cellAddress]) {
        const cellValue = ws[cellAddress].v;
        // 检查是否是空列的列名
        if (cellValue === undefined ||
            (typeof cellValue === 'string' && cellValue.startsWith('__empty_')) ||
            cellValue === '') {
          ws[cellAddress].v = ''; // 设置为空字符串
        }
      }
    }

    // 获取第一行的所有列名
    const headerColumns = [];
    for (let col = range.s.c; col <= range.e.c; col++) {
      const cellAddress = XLSX.utils.encode_cell({ r: 0, c: col });
      if (ws[cellAddress]) {
        headerColumns.push(ws[cellAddress].v);
      }
    }

    // 遍历所有单元格，添加边框和样式
    for (let row = range.s.r; row <= range.e.r; row++) {
      for (let col = range.s.c; col <= range.e.c; col++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
        if (!ws[cellAddress]) continue;

        const columnName = headerColumns[col];

        // 空列：不设置任何样式（无边框、无背景、无文字）
        if (columnName === '' || columnName === undefined || (typeof columnName === 'string' && columnName.startsWith('__empty_'))) {
          continue; // 直接跳过空列，不添加任何样式
        }

        // 检查是否是部门或延期未发布数量列
        const isStatsColumn = columnName === '部门' || columnName === '延期未发布数量';

        // 对于统计列，只有单元格有值时才设置样式
        if (isStatsColumn && row > 0) {
          const cellValue = ws[cellAddress].v;
          if (cellValue === undefined || cellValue === null || cellValue === '' || cellValue === 0) {
            continue; // 空白单元格跳过，不设置任何样式
          }
        }

        // 设置基础样式（边框）
        const cellStyle = {
          border: {
            top: { style: 'thin', color: { rgb: '000000' } },
            bottom: { style: 'thin', color: { rgb: '000000' } },
            left: { style: 'thin', color: { rgb: '000000' } },
            right: { style: 'thin', color: { rgb: '000000' } }
          }
        };

        // 第一行（标题行）添加背景色
        if (row === 0) {
          cellStyle.fill = {
            fgColor: { rgb: '4472C4' } // 蓝色背景
          };
          cellStyle.font = {
            color: { rgb: 'FFFFFF' }, // 白色字体
            bold: true
          };
        }

        ws[cellAddress].s = cellStyle;
      }
    }

    // 自动设置统一列宽
    const colCount = range.e.c - range.s.c + 1;
    const colWidths = [];

    // 计算所有单元格的平均宽度
    let totalWidth = 0;
    let cellCount = 0;

    for (let col = range.s.c; col <= range.e.c; col++) {
      for (let row = range.s.r; row <= range.e.r; row++) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: col });
        const cell = ws[cellAddress];

        if (cell && cell.v) {
          let cellWidth = 0;
          const cellValue = String(cell.v);
          for (let i = 0; i < cellValue.length; i++) {
            const char = cellValue.charCodeAt(i);
            if (char > 255) {
              cellWidth += 2;
            } else {
              cellWidth += 1;
            }
          }
          totalWidth += cellWidth;
          cellCount++;
        }
      }
    }

    // 计算平均宽度，并设置合理的统一宽度（最小12，最大25）
    const avgWidth = cellCount > 0 ? totalWidth / cellCount : 15;
    const uniformWidth = Math.min(Math.max(Math.round(avgWidth + 2), 12), 25);

    // 所有列使用统一宽度
    for (let i = 0; i < colCount; i++) {
      colWidths.push({ wch: uniformWidth });
    }

    // 应用列宽设置
    ws['!cols'] = colWidths;

    // 添加总计（在部门列和延期未发布数量列的最后一个有数据的行）
    // 找到部门列和延期未发布数量列的索引
    let deptColIndex = -1;
    let countColIndex = -1;
    for (let i = 0; i < headerColumns.length; i++) {
      if (headerColumns[i] === '部门') {
        deptColIndex = i;
      } else if (headerColumns[i] === '延期未发布数量') {
        countColIndex = i;
      }
    }

    // 如果找到了这两列，添加总计
    if (deptColIndex >= 0 && countColIndex >= 0) {
      // 计算延期未发布数量列的总和
      let totalSum = 0;
      for (let row = 1; row <= range.e.r; row++) { // 从第2行开始（跳过标题行）
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: countColIndex });
        const cell = ws[cellAddress];
        if (cell && cell.v !== undefined && cell.v !== '') {
          const value = parseFloat(cell.v);
          if (!isNaN(value)) {
            totalSum += value;
          }
        }
      }

      // 找到部门列最后一个有值的行
      let lastDataRow = range.e.r;
      for (let row = range.e.r; row >= 1; row--) {
        const cellAddress = XLSX.utils.encode_cell({ r: row, c: deptColIndex });
        const cell = ws[cellAddress];
        if (cell && cell.v !== undefined && cell.v !== '' && cell.v !== 0) {
          lastDataRow = row;
          break;
        }
      }

      // 在最后一个有数据的行设置总计
      const summaryCellAddress = XLSX.utils.encode_cell({ r: lastDataRow, c: deptColIndex });
      ws[summaryCellAddress] = {
        v: '总计',
        t: 's',
        s: {
          border: {
            top: { style: 'thin', color: { rgb: '000000' } },
            bottom: { style: 'thin', color: { rgb: '000000' } },
            left: { style: 'thin', color: { rgb: '000000' } },
            right: { style: 'thin', color: { rgb: '000000' } }
          },
          fill: {
            fgColor: { rgb: '4472C4' } // 蓝色背景
          },
          font: {
            bold: true,
            color: { rgb: 'FFFFFF' } // 白色字体
          }
        }
      };

      // 设置总和单元格（延期未发布数量列）
      const sumCellAddress = XLSX.utils.encode_cell({ r: lastDataRow, c: countColIndex });
      ws[sumCellAddress] = {
        v: totalSum,
        t: 'n',
        s: {
          border: {
            top: { style: 'thin', color: { rgb: '000000' } },
            bottom: { style: 'thin', color: { rgb: '000000' } },
            left: { style: 'thin', color: { rgb: '000000' } },
            right: { style: 'thin', color: { rgb: '000000' } }
          }
        }
      };
    }

    // 将工作表添加到工作簿
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

    // 导出文件
    const today = new Date();
    const year = today.getFullYear();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    const day = String(today.getDate()).padStart(2, '0');
    const dateStr = `${year}${month}${day}`;
    XLSX.writeFile(wb, `延期未发布ESO零件-${dateStr}.xlsx`);

    ElMessage.success('导出成功！');
  } catch (error) {
    console.error('Export error:', error);
    ElMessage.error('导出失败：' + error.message);
  }
}

// 从 localStorage 读取保存的列选择
const getSavedColumnSelection = () => {
  try {
    const saved = localStorage.getItem('chatbot_column_selection');
    return saved ? JSON.parse(saved) : null;
  } catch (e) {
    console.error('读取列选择失败:', e);
    return null;
  }
};

// 获取初始列选择（优先使用保存的，否则使用默认值）
const getInitialColumnSelection = (allColumns, defaultColumns) => {
  const saved = getSavedColumnSelection();

  // 如果没有保存的选择，使用默认值
  if (!saved || saved.length === 0) {
    return defaultColumns;
  }

  // 过滤出当前表格中实际存在的列
  const validSelection = saved.filter(col => allColumns.includes(col));

  // 如果保存的选择都不在当前表格中，使用默认值
  if (validSelection.length === 0) {
    return defaultColumns;
  }

  return validSelection;
};

// 保存列选择到 localStorage
const saveColumnSelection = (selectedColumns) => {
  try {
    localStorage.setItem('chatbot_column_selection', JSON.stringify(selectedColumns));
  } catch (e) {
    console.error('保存列选择失败:', e);
  }
};

// 处理列全选/全不选
const toggleAllColumns = (checked, tableData) => {
  if (checked) {
    tableData.selectedColumns = tableData.columns.map(col => col.prop);
    tableData.previousSelectedColumns = [...tableData.selectedColumns];
  } else {
    tableData.selectedColumns = [];
    tableData.previousSelectedColumns = [];
  }
  // 保存选择到 localStorage
  saveColumnSelection(tableData.selectedColumns);
}

// 处理单个列的选择变化，新选择的列插入到ESO_Actual_Date左边
const handleColumnChange = (selectedColumns, tableData) => {
  // 统计列的名称
  const statsColumns = ['部门', '延期未发布数量'];

  // 获取上一次的选择状态
  const previousSelected = tableData.previousSelectedColumns || [];

  // 找出新增的列（在previousSelected中不存在，但在selectedColumns中存在）
  const newColumns = selectedColumns.filter(col => !previousSelected.includes(col));

  // 找出被取消选择的列（在previousSelected中存在，但在selectedColumns中不存在）
  const removedColumns = previousSelected.filter(col => !selectedColumns.includes(col));

  console.log('[handleColumnChange] 新增列:', newColumns, '取消列:', removedColumns);

  // 分离统计列
  const statsColsSelected = selectedColumns.filter(col => statsColumns.includes(col));

  // 将列分为几部分：
  // 1. ESO_Actual_Date及其左边的列（保持原顺序）
  // 2. 新增的列（放在ESO_Actual_Date左边）
  // 3. ESO_Actual_Date右边的非统计列
  // 4. 统计列

  const esoActualDateCol = 'ESO_Actual_Date';
  const otherColumns = selectedColumns.filter(col => !statsColumns.includes(col));

  // 找到ESO_Actual_Date的位置（基于当前选择）
  const esoActualDateIndex = otherColumns.indexOf(esoActualDateCol);

  let orderedColumns;

  if (esoActualDateIndex >= 0 && newColumns.length > 0) {
    // 如果ESO_Actual_Date存在且有新增列，将新增列插入到它左边
    const leftColumns = otherColumns.slice(0, esoActualDateIndex);
    const rightColumns = otherColumns.slice(esoActualDateIndex + 1);

    // 从leftColumns和rightColumns中移除新增的列（避免重复）
    const filteredLeftColumns = leftColumns.filter(col => !newColumns.includes(col));
    const filteredRightColumns = rightColumns.filter(col => !newColumns.includes(col));

    // 新增列插入到ESO_Actual_Date左边
    orderedColumns = [...filteredLeftColumns, ...newColumns, esoActualDateCol, ...filteredRightColumns];
  } else {
    // 如果ESO_Actual_Date不存在或没有新增列，保持原顺序
    orderedColumns = otherColumns;
  }

  // 最后添加统计列
  const finalOrderedColumns = [...orderedColumns, ...statsColsSelected];

  // 更新selectedColumns和previousSelectedColumns
  tableData.previousSelectedColumns = [...finalOrderedColumns];
  tableData.selectedColumns = finalOrderedColumns;

  // 保存选择到 localStorage
  saveColumnSelection(tableData.selectedColumns);

  console.log('[handleColumnChange] 最终顺序:', tableData.selectedColumns);
}

// 获取按选择顺序排序的列
const getOrderedColumns = (tableData) => {
  if (tableData.exportAllColumns) {
    return tableData.columns;
  }

  // 按照selectedColumns的顺序返回列
  return tableData.selectedColumns
    .map(prop => tableData.columns.find(col => col.prop === prop))
    .filter(col => col); // 过滤掉undefined
}

// 处理图像加载失败
const handleImageError = (event) => {
  event.target.src = '/images/placeholder.png'; // 设置默认占位图
}

// 显示图像预览
const showImagePreview = (imageUrl) => {
  imageViewerUrl.value = imageUrl;
  showImageViewer.value = true;
}

// 解析可能的表格数据
const parseTableData = (content) => {
  try {
    // 首先尝试解析整个内容
    let parsedContent = JSON.parse(content);
    console.log('Parsed content:', parsedContent);

    // 检查是否包含 json 字段
    let data = parsedContent;
    if (data.hasOwnProperty('json')) {
      data = data.json;
    }

    // 检查是否有统计信息 - stats_info或group_stats可能在最外层(data)或在parsedContent中
    let statsData = [];
    let hasStatsInfo = false;

    // 优先检查 group_stats (UPDATE操作后的统计)
    if (parsedContent.hasOwnProperty('group_stats') && parsedContent.group_stats.hasOwnProperty('group_data')) {
      statsData = parsedContent.group_stats.group_data;
      hasStatsInfo = true;
      console.log('Found group_stats in parsedContent, statsData:', statsData);
    } else if (data.hasOwnProperty('group_stats') && data.group_stats.hasOwnProperty('group_data')) {
      statsData = data.group_stats.group_data;
      hasStatsInfo = true;
      console.log('Found group_stats in data, statsData:', statsData);
    }
    // 其次检查 stats_info (SELECT操作时的统计)
    else if (parsedContent.hasOwnProperty('stats_info') && parsedContent.stats_info.hasOwnProperty('stats_data')) {
      statsData = parsedContent.stats_info.stats_data;
      hasStatsInfo = true;
      console.log('Found stats_info in parsedContent, statsData:', statsData);
    } else if (data.hasOwnProperty('stats_info') && data.stats_info.hasOwnProperty('stats_data')) {
      statsData = data.stats_info.stats_data;
      hasStatsInfo = true;
      console.log('Found stats_info in data, statsData:', statsData);
    } else {
      console.log('No stats_info or group_stats found in parsedContent or data');
    }

    // 不需要字段映射，直接使用原始字段名
    console.log('Final statsData:', statsData);

    // 检查data数组中是否包含带"数据类型"字段的统计行
    let dataRows = []
    let embeddedStatsRows = []

    // 提取数据行和统计行
    if (Array.isArray(data)) {
      data.forEach(item => {
        if (typeof item === 'object' && item !== null) {
          if (item.hasOwnProperty('数据类型') && item['数据类型'] === '统计结果') {
            embeddedStatsRows.push(item)
          } else {
            dataRows.push(item)
          }
        } else {
          dataRows.push(item)
        }
      })
    }

    // 如果找到了嵌入的统计行,并且没有独立的stats_info,则使用嵌入的统计行
    if (embeddedStatsRows.length > 0 && !hasStatsInfo) {
      console.log('Found embedded stats rows in data, count:', embeddedStatsRows.length)
      // 转换统计行格式
      statsData = embeddedStatsRows.map(row => ({
        '部门': row['部门'],
        'count': row['count']
      }))
      data = dataRows // 使用过滤后的数据行
    }

    console.log('Final statsData:', statsData)
    console.log('Final data:', data)

    // 检查是否是表格数据格式 - 处理 [{records: [...]}] 格式
    if (Array.isArray(data) && data.length > 0 && data[0].hasOwnProperty('records')) {
      let records = data[0].records
      const total = records.length

      console.log('[{records}格式] 原始records长度:', total)
      console.log('[{records}格式] 原始records最后3条:', records.slice(-3))
      console.log('[{records}格式] 第1条记录的keys:', Object.keys(records[0] || {}))

      // 首先检查records数组中是否包含带"数据类型"字段的统计行
      let dataRows = []
      let embeddedStatsRows = []

      records.forEach(item => {
        if (typeof item === 'object' && item !== null) {
          if (item.hasOwnProperty('数据类型') && item['数据类型'] === '延期未发布数量') {
            embeddedStatsRows.push(item)
          } else {
            dataRows.push(item)
          }
        } else {
          dataRows.push(item)
        }
      })

      console.log('[{records}格式] 提取结果 - dataRows:', dataRows.length, 'embeddedStatsRows:', embeddedStatsRows.length)

      // 如果找到了嵌入的统计行,并且没有独立的stats_info,则使用嵌入的统计行
      if (embeddedStatsRows.length > 0 && statsData.length === 0) {
        console.log('[{records}格式] Found embedded stats rows in records, count:', embeddedStatsRows.length)
        console.log('[{records}格式] embeddedStatsRows前3条:', embeddedStatsRows.slice(0, 3))
        // 转换统计行格式
        statsData = embeddedStatsRows.map(row => ({
          '部门': row['部门'],
          '延期未发布数量': row['延期未发布数量']
        }))
        records = dataRows // 使用过滤后的数据行
        console.log('[{records}格式] 提取后的statsData:', statsData)
      }

      if (total > 0 && typeof records[0] === 'object' && records[0] !== null) {
        // 获取列信息，设置适中的列宽
        let columns = Object.keys(records[0]).map(key => ({
          prop: key,
          label: key,
          width: Math.min(Math.max(key.length * 8, 90), 200) // 调整为适中的宽度范围：90-200px
        }))

        // 如果有统计信息，将统计结果插入到records开头
        if (statsData.length > 0) {
          console.log('[{records}格式] 准备插入统计行,统计数据条数:', statsData.length)

          // 确保"功能组统计"和"统计数量"列存在
          if (!columns.find(c => c.prop === '功能组统计')) {
            columns.push({ prop: '部门', label: '部门', width: 120 })
          }
          if (!columns.find(c => c.prop === '统计数量')) {
            columns.push({ prop: '延期未发布数量', label: '延期未发布数量', width: 150 })
          }
          console.log('[{records}格式] 添加统计列后的columns:', columns.map(c => c.prop))
          // 创建统计行对象数组
          const statRows = []
          statsData.forEach(statItem => {
            const statRow = {}
            // 设置功能组统计列
            statRow['部门'] = statItem['部门']
            statRow['延期未发布数量'] = statItem['延期未发布数量']
            // 其他列设置为空字符串
            columns.forEach(col => {
              if (col.prop !== '部门' && col.prop !== '延期未发布数量') {
                statRow[col.prop] = ''
              }
            })
            statRows.push(statRow)
          })
          console.log('[{records}格式] 创建的统计行:', statRows)
          console.log('[{records}格式] 插入前records长度:', records.length)
          // 将统计行插入到records的最前面
          records.unshift(...statRows)
          console.log('[{records}格式] 插入后records长度:', records.length)
          console.log('[{records}格式] 插入后records前3条:', records.slice(0, 3))
        } else {
          console.log('[{records}格式] statsData为空,无法插入统计行')
        }

        // 重新排序列，将特定列放在前面
        const fixedOrder = ['零件号', 'FFC中文描述', '功能组', '工程师', 'ESO_Plan_Date', 'ESO材料送检计划', 'ESO备注_类型', '是否需要ESO_物流提供_', 'ESO_Actual_Date', '部门', '延期未发布数量', 'ESO备注', '首次申请项目'];
        const fixedColumns = [];
        const otherColumns = [];

        columns.forEach(col => {
          if (fixedOrder.includes(col.label)) {
            fixedColumns.push(col);
          } else {
            otherColumns.push(col);
          }
        });

        // 按照固定的顺序排列固定列
        const orderedFixedColumns = [];
        fixedOrder.forEach(label => {
          const col = fixedColumns.find(c => c.label === label);
          if (col) {
            orderedFixedColumns.push(col);
          }
        });

        // 合并列数组：固定列 + 其他列
        columns = [...orderedFixedColumns, ...otherColumns];

        
        // 默认显示前10条记录
        const pageSize = 10
        const displayRecords = records.slice(0, pageSize)
        
        return {
          records, // 所有记录
          displayRecords, // 当前显示的记录
          columns,
          total,
          pageSize,
          currentPage: 1,
          pagination: total > pageSize,
          // 添加列选择状态，优先使用保存的选择
          selectedColumns: getInitialColumnSelection(
            columns.map(c => c.prop),
            orderedFixedColumns.map(col => col.prop)
          ),
          previousSelectedColumns: getInitialColumnSelection(
            columns.map(c => c.prop),
            orderedFixedColumns.map(col => col.prop)
          ), // 用于跟踪之前的选择状态
          exportAllColumns: false // 默认不导出所有列
        }
      }
    }
    // 处理 {records: [...]} 格式
    else if (typeof data === 'object' && data !== null && data.hasOwnProperty('records') && Array.isArray(data.records)) {
      const records = data.records
      const total = records.length

      if (total > 0 && typeof records[0] === 'object' && records[0] !== null) {
        // 获取列信息，设置适中的列宽
        let columns = Object.keys(records[0]).map(key => ({
          prop: key,
          label: key,
          width: Math.min(Math.max(key.length * 8, 90), 200) // 调整为适中的宽度范围：90-200px
        }))

        // 如果有统计信息，将统计结果插入到records开头
        if (statsData.length > 0) {
          // 确保"功能组统计"和"统计数量"列存在
          if (!columns.find(c => c.prop === '部门')) {
            columns.push({ prop: '部门', label: '部门', width: 120 })
          }
          if (!columns.find(c => c.prop === '延期未发布数量')) {
            columns.push({ prop: '延期未发布数量', label: '延期未发布数量', width: 150 })
          }

          // 创建统计行对象数组
          const statRows = []
          statsData.forEach(statItem => {
            const statRow = {}
            // 设置功能组统计列
            statRow['部门'] = statItem['部门']
            statRow['延期未发布数量'] = statItem['延期未发布数量']
            // 其他列设置为空字符串
            columns.forEach(col => {
              if (col.prop !== '部门' && col.prop !== '延期未发布数量') {
                statRow[col.prop] = ''
              }
            })
            statRows.push(statRow)
          })
          // 将统计行插入到records的最前面
          records.unshift(...statRows)
        }

        // 重新排序列，将特定列放在前面
        const fixedOrder = ['零件号', 'FFC中文描述', '功能组', '工程师', 'ESO_Plan_Date', 'ESO材料送检计划', 'ESO备注_类型', '是否需要ESO_物流提供_', 'ESO_Actual_Date', '部门', '延期未发布数量', 'ESO备注', '首次申请项目'];
        const fixedColumns = [];
        const otherColumns = [];

        columns.forEach(col => {
          if (fixedOrder.includes(col.label)) {
            fixedColumns.push(col);
          } else {
            otherColumns.push(col);
          }
        });

        // 按照固定的顺序排列固定列
        const orderedFixedColumns = [];
        fixedOrder.forEach(label => {
          const col = fixedColumns.find(c => c.label === label);
          if (col) {
            orderedFixedColumns.push(col);
          }
        });

        // 合并列数组：固定列 + 其他列
        columns = [...orderedFixedColumns, ...otherColumns];

        
        // 默认显示前10条记录
        const pageSize = 10
        const displayRecords = records.slice(0, pageSize)
        
        return {
          records, // 所有记录
          displayRecords, // 当前显示的记录
          columns,
          total,
          pageSize,
          currentPage: 1,
          pagination: total > pageSize,
          // 添加列选择状态，优先使用保存的选择
          selectedColumns: getInitialColumnSelection(
            columns.map(c => c.prop),
            orderedFixedColumns.map(col => col.prop)
          ),
          previousSelectedColumns: getInitialColumnSelection(
            columns.map(c => c.prop),
            orderedFixedColumns.map(col => col.prop)
          ), // 用于跟踪之前的选择状态
          exportAllColumns: false // 默认不导出所有列
        }
      }
    }
    // 处理普通数组数据
    else if (Array.isArray(data) && data.length > 0 && typeof data[0] === 'object' && data[0] !== null) {
      const total = data.length

      // 如果有统计信息，将统计结果插入到data开头
      if (statsData.length > 0) {
        console.log('[普通数组格式] 准备插入统计行,统计数据条数:', statsData.length)
        console.log('[普通数组格式] 原始data前3条:', data.slice(0, 3))

        // 获取原始数据的列(排除统计行)
        let dataColumns = Object.keys(data[0] || {})

        // 创建统计行对象数组
        const statRows = []
        statsData.forEach(statItem => {
          const statRow = {}
          // 设置功能组统计列
          statRow['部门'] = statItem['部门']
          statRow['延期未发布数量'] = statItem['延期未发布数量']
          // 其他列设置为空字符串
          dataColumns.forEach(col => {
            if (col !== '部门' && col !== '延期未发布数量') {
              statRow[col] = ''
            }
          })
          statRows.push(statRow)
        })

        console.log('[普通数组格式] 创建的统计行:', statRows)
        console.log('[普通数组格式] 插入前data长度:', data.length)

        // 将统计行插入到data的最前面
        data.unshift(...statRows)

        console.log('[普通数组格式] 插入后data长度:', data.length)
        console.log('[普通数组格式] 插入后data前3条:', data.slice(0, 3))
      }

      let columns = Object.keys(data[0]).map(key => ({
        prop: key,
        label: key,
        width: Math.min(Math.max(key.length * 8, 90), 200) // 调整为适中的宽度范围：90-200px
      }))

      // 重新排序列，将特定列放在前面
      const fixedOrder = ['零件号', 'FFC中文描述', '功能组', '工程师', 'ESO_Plan_Date', 'ESO材料送检计划', 'ESO备注_类型', '是否需要ESO_物流提供_', 'ESO_Actual_Date', '部门', '延期未发布数量', 'ESO备注', '首次申请项目'];
      const fixedColumns = [];
      const otherColumns = [];

      columns.forEach(col => {
        if (fixedOrder.includes(col.label)) {
          fixedColumns.push(col);
        } else {
          otherColumns.push(col);
        }
      });

      // 按照固定的顺序排列固定列
      const orderedFixedColumns = [];
      fixedOrder.forEach(label => {
        const col = fixedColumns.find(c => c.label === label);
        if (col) {
          orderedFixedColumns.push(col);
        }
      });

      // 合并列数组：固定列 + 其他列
      columns = [...orderedFixedColumns, ...otherColumns];

      
      // 默认显示前10条记录
      const pageSize = 10
      const displayRecords = data.slice(0, pageSize)
      
      return {
        records: data, // 所有记录
        displayRecords, // 当前显示的记录
        columns,
        total,
        pageSize,
        currentPage: 1,
        pagination: total > pageSize,
        // 添加列选择状态，默认只选择前8列
        selectedColumns: orderedFixedColumns.map(col => col.prop), // 默认只选择指定的8列
        previousSelectedColumns: [...orderedFixedColumns.map(col => col.prop)], // 用于跟踪之前的选择状态
        exportAllColumns: false // 默认不导出所有列
      }
    }
    return null
  } catch (e) {
    // 不是有效的JSON数据
    console.log('Failed to parse table data:', e)
    return null
  }
}

const sendMessage = async () => {
  if (!inputMessage.value.trim() || sending.value) return

  const userMessage = inputMessage.value.trim()

  // 添加用户消息到聊天记录
  messages.value.push({
    type: 'sent',
    sender: '您',
    text: userMessage,
    timestamp: new Date().toLocaleTimeString()
  })

  // 清空输入框并滚动到底部
  inputMessage.value = ''
  scrollToBottom()

  try {
    sending.value = true

    // 发送到Dify Workflow API
    const response = await runWorkflow(userMessage)
    console.log('Received response:', response)

    // 尝试解析表格数据
    const tableData = parseTableData(response.content)

    // 如果表格数据来自sheet表,额外获取功能组统计数据
    if (tableData && tableData.total > 0) {
      try {
        // 检查是否有功能组列
        const hasFunctionGroup = tableData.records.some(record =>
          record.hasOwnProperty('功能组') || record.hasOwnProperty('部门')
        )

        if (hasFunctionGroup) {
          let statsData = null

          // 首先检查响应中是否已经包含group_stats
          try {
            const responseData = JSON.parse(response.content)
            if (responseData.group_stats && responseData.group_stats.group_data) {
              statsData = responseData.group_stats.group_data
            }
            // 如果没有group_stats,检查是否有ss_info,从ss_data中统计
            else if (responseData.ss_info && responseData.ss_info.ss_data) {
              const ssData = responseData.ss_info.ss_data
              const groupStats = {}
              for (const row of ssData) {
                const 功能组 = row['功能组']
                if (功能组) {
                  groupStats[功能组] = (groupStats[功能组] || 0) + 1
                }
              }
              statsData = Object.entries(groupStats).map(([k, v]) => ({
                '部门': k,
                '延期未发布数量': v
              }))
            }
          } catch (e) {
            console.log('解析响应数据失败:', e)
          }

          // 如果还是没有统计数据,则从tableData的records中统计功能组
          if (!statsData) {
            const groupStats = {}
            for (const row of tableData.records) {
              const 功能组 = row['功能组']
              if (功能组) {
                groupStats[功能组] = (groupStats[功能组] || 0) + 1
              }
            }
            statsData = Object.entries(groupStats).map(([k, v]) => ({
              '部门': k,
              '延期未发布数量': v
            }))
          }

          // 添加统计列到表格数据
          if (statsData && statsData.length > 0) {
            // 为前N行添加统计列(如果记录数少于统计条数,则为所有行)
            const statsRowCount = Math.min(statsData.length, tableData.records.length)

            for (let i = 0; i < statsRowCount; i++) {
              tableData.records[i]['部门'] = statsData[i]['部门']
              tableData.records[i]['延期未发布数量'] = statsData[i]['延期未发布数量']
            }

            // 更新列定义，确保包含"功能组统计"和"统计数量"列
            if (!tableData.columns.find(c => c.prop === '部门')) {
              tableData.columns.push({ prop: '部门', label: '部门', width: 120 })
            }
            if (!tableData.columns.find(c => c.prop === '延期未发布数量')) {
              tableData.columns.push({ prop: '延期未发布数量', label: '延期未发布数量', width: 150 })
            }
          }
        }
      } catch (statsError) {
        console.error('处理统计数据失败:', statsError)
        // 不影响主流程,继续显示表格数据
      }
    }

    if (tableData && tableData.total > 0) {
      // 添加带表格的AI回复到聊天记录
      messages.value.push({
        type: 'received',
        sender: 'AI助手',
        tableData: tableData,
        timestamp: new Date().toLocaleTimeString()
      })
    } else {
      // 添加普通文本AI回复到聊天记录
      const answer = response.content || '抱歉，我没有理解您的问题。'
      messages.value.push({
        type: 'received',
        sender: 'AI助手',
        text: answer,
        timestamp: new Date().toLocaleTimeString()
      })
    }
  } catch (error) {
    console.error('发送消息失败:', error)
    ElMessage.error('发送消息失败: ' + error.message)

    // 添加错误消息到聊天记录
    messages.value.push({
      type: 'received',
      sender: '系统',
      text: '抱歉，发送消息时出现错误，请稍后重试。',
      timestamp: new Date().toLocaleTimeString()
    })
  } finally {
    sending.value = false
    scrollToBottom()
  }
}

// 不需要单独的close方法，el-image-viewer会自动处理

</script>

<style scoped>
.chat-container-wrapper {
  height: 100%;
  display: flex;
  flex-direction: column;
}

.chat-container {
  display: flex;
  flex-direction: column;
  height: 100%;
  width: 100%;
  max-width: 100%;
  margin: 0 auto;
  border: 1px solid #e4e7ed;
  border-radius: 8px;
  overflow: hidden;
  box-shadow: 0 2px 12px 0 rgba(0, 0, 0, 0.1);
  flex: 1;
}

.chat-history {
  flex: 1;
  overflow-y: auto;
  padding: 20px;
  background-color: #f5f7fa;
  height: calc(100% - 120px); /* 减去输入区域的大致高度 */
}

.message {
  margin-bottom: 15px;
  display: flex;
}

.message.sent {
  justify-content: flex-end;
}

.message.received {
  justify-content: flex-start;
}

.message-content {
  max-width: 80%;
  padding: 10px 15px;
  border-radius: 10px;
  position: relative;
}

.message.sent .message-content {
  background-color: #409eff;
  color: white;
}

.message.received .message-content {
  background-color: white;
  border: 1px solid #e4e7ed;
}

.sender {
  font-weight: bold;
  margin-bottom: 5px;
  font-size: 14px;
}

.text {
  font-size: 15px;
  line-height: 1.5;
  margin-bottom: 5px;
  white-space: pre-wrap;
}

.timestamp {
  font-size: 12px;
  opacity: 0.7;
  text-align: right;
}

.input-area {
  padding: 20px;
  background-color: white;
  border-top: 1px solid #e4e7ed;
}

.send-button {
  margin-top: 10px;
  float: right;
}

.table-container {
  margin: 10px 0;
}

.table-summary {
  margin-bottom: 10px;
  font-weight: bold;
  color: #409eff;
}

.export-button-container {
  margin-bottom: 10px;
  text-align: right;
}

.column-selector {
  margin: 10px 0;
  display: flex;
  gap: 10px;
}

.column-options {
  max-height: 300px;
  overflow-y: auto;
}

.column-checkbox-group {
  display: flex;
  flex-direction: column;
  margin-top: 10px;
}

.column-checkbox {
  margin-bottom: 8px;
  display: flex;
  align-items: center;
}

.pagination-container {
  margin-top: 15px;
  display: flex;
  justify-content: center;
}

/* 减小表格行高 */
:deep(.el-table__row) {
  height: 30px;
}

:deep(.el-table__row td) {
  padding: 2px 0;
}

:deep(.el-table th) {
  padding: 4px 0;
}

.examples-container {
  margin-top: 15px;
}

.example-item {
  margin-bottom: 15px;
  padding: 10px;
  border: 1px solid #ebeef5;
  border-radius: 4px;
  background-color: #f5f7fa;
}

.example-question {
  margin-bottom: 10px;
  padding-bottom: 5px;
  border-bottom: 1px dashed #dcdfe6;
}

.example-image {
  text-align: center;
  margin: 10px 0;
}

.example-image img {
  max-width: 100%;
  max-height: 200px;
  border: 1px solid #dcdfe6;
  border-radius: 4px;
}

.clickable-image {
  cursor: pointer;
  transition: all 0.3s ease;
}

.clickable-image:hover {
  transform: scale(1.03);
  box-shadow: 0 4px 12px rgba(0, 0, 0, 0.15);
}

.clickable-image {
  cursor: pointer;
}

.example-placeholder {
  text-align: center;
  padding: 20px;
  color: #909399;
  background-color: #f4f4f5;
  border-radius: 4px;
  margin: 10px 0;
}
</style>
