import axios from 'axios'

const apiKey = import.meta.env.VITE_DIFY_API_KEY
if (!apiKey) {
  console.warn('警告: VITE_DIFY_API_KEY 未设置')
}

const difyClient = axios.create({
  baseURL: 'http://10.182.38.53/v1', // 替换为你的 Dify 地址
  headers: {
    'Content-Type': 'application/json',
    'Authorization': `Bearer ${apiKey}`
  }
})

/**
 * 调用 Dify Workflow（工作流模式）
 * @param {string} query - 用户的自然语言查询
 * @returns {Promise<{ content: string, rawData: any }>}
 */
export const runWorkflow = async (query) => {
  // 构造工作流输入参数 - 根据错误信息，需要正确的inputs格式
  const payload = {
    inputs: {
      inputs: query
    },
    response_mode: 'blocking',
    user: 'engineering-data-user'
  }

  try {
    // 模拟一些示例数据用于测试
    if (query.includes('示例') || query.includes('example')) {
      await new Promise(resolve => setTimeout(resolve, 1000))
      return {
        content: JSON.stringify(generateSampleData(query)),
        rawData: {}
      }
    }

    console.log('发送到Dify的请求:', payload)
    
    // 调用工作流接口
    const res = await difyClient.post('/workflows/run', payload)
    
    console.log('Dify Workflow 响应:', res.data)

    // Workflow 返回结构：res.data.data.outputs
    const outputs = res.data?.data?.outputs || {}

    // 根据实际响应结构调整解析方式
    let content = ''
    if (outputs.result) {
      content = outputs.result
    } else if (Object.keys(outputs).length > 0) {
      // 如果没有result字段，但有其他输出字段，就序列化整个outputs
      content = JSON.stringify(outputs)
    } else {
      // fallback到原始响应
      content = JSON.stringify(res.data)
    }

    return {
      content,
      rawData: res.data
    }
  } catch (error) {
    console.error('Dify Workflow 请求失败:', {
      status: error.response?.status,
      data: error.response?.data,
      message: error.message
    })

    if (error.response?.status === 400) {
      throw new Error('请求参数错误: ' + (error.response.data?.message || ''))
    }
    if (error.response?.status === 401 || error.response?.status === 403) {
      throw new Error('API 密钥无效或无权限')
    }
    if (error.code === 'ECONNREFUSED' || error.code === 'ENOTFOUND') {
      throw new Error('无法连接到 Dify 服务，请检查网络和地址')
    }

    throw new Error(`工作流执行失败: ${error.message}`)
  }
}

// 生成示例数据的函数
function generateSampleData(query) {
  // 如果查询中包含"表格"或"table"，返回表格数据
  if (query.includes('表格') || query.includes('table') || query.includes('记录')) {
    return [
      { vin: 'LSJE36098NS086649', total_power_on_hours: 1.13, mileage: 120.5, status: '运行中' },
      { vin: 'LSJE3609XNS090069', total_power_on_hours: 0.51, mileage: 85.2, status: '停车中' },
      { vin: 'LSJWL4090SS035456', total_power_on_hours: 0.09, mileage: 45.7, status: '维修中' },
      { vin: 'LSJWL4092SS035457', total_power_on_hours: 0.11, mileage: 62.3, status: '运行中' },
      { vin: 'LSJWL4092SS035460', total_power_on_hours: 0.12, mileage: 78.9, status: '停车中' }
    ]
  }
  
  return {
    message: "这是一条示例消息",
    data: [
      { name: "车辆A", status: "运行中", location: "北京" },
      { name: "车辆B", status: "停车中", location: "上海" }
    ]
  }
}