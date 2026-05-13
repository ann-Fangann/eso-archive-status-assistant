<template>
  <div class="upload-container">
    <el-upload
      action="/api/upload-excel/"
      :auto-upload="false"
      :on-change="handleChange"
      :on-success="handleSuccess"
      :on-error="handleError"
      accept=".xlsx,.xls"
      drag
      :limit="1"
      :file-list="fileList"
    >
      <i class="el-icon-upload"></i>
      <div class="el-upload__text">将文件拖到此处，或点击上传</div>
      <div class="el-upload__tip" slot="tip">只能上传xlsx/xls文件</div>
    </el-upload>
    
    <div v-if="result" class="result-container">
      <h3>上传结果</h3>
      <div v-if="result.success">
        <p>文件上传成功！</p>
        <h4>字段说明：</h4>
        <ul>
          <li v-for="(desc, field) in result.field_descriptions" :key="field">
            <strong>{{ field }}:</strong> {{ desc }}
          </li>
        </ul>
        <p>数据预览：</p>
        <el-table :data="result.data_preview" style="width: 100%">
          <el-table-column v-for="(col, index) in Object.keys(result.data_preview[0])" :prop="col" :label="col" :key="index"></el-table-column>
        </el-table>
        <p>总行数：{{ result.total_rows }}</p>
      </div>
      <div v-else>
        <p style="color: red;">上传失败：{{ result.message }}</p>
      </div>
    </div>
  </div>
</template>

<script>
export default {
  data() {
    return {
      fileList: [],
      result: null
    }
  },
  methods: {
    handleChange(file, fileList) {
      this.fileList = fileList
    },
    async handleSuccess(response, file, fileList) {
      this.result = response
    },
    handleError(err, file, fileList) {
      this.result = { success: false, message: err.response.data.detail }
    }
  }
}
</script>