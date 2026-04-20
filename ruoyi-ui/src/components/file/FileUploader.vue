<template>
    <div class="file-uploader">
        <el-upload class="upload-container"
                   action="#"
                   :auto-upload="false"
                   :on-change="handleFileSelect"
                   :file-list="fileList"
                   :before-upload="beforeFileUpload"
                   :accept="'.pdf,.doc,.docx,.xls,.xlsx,.txt,.md'"
                   multiple>
            <el-button size="mini" type="default" icon="el-icon-upload2">选择文件</el-button>
        </el-upload>
        <div class="file-list">
            <div v-for="(file, index) in fileList" :key="index" class="file-item">
                <span class="file-name">{{ file.name }}</span>
                <span class="file-size">{{ formatFileSize(file.size) }}</span>
                <el-button size="mini"
                           type="text"
                           icon="el-icon-delete"
                           @click="removeFile(index)"
                           class="delete-btn"></el-button>
            </div>
        </div>
        <el-button size="mini"
                   type="primary"
                   icon="el-icon-upload"
                   class="upload-btn"
                   @click="handleUpload"
                   :disabled="fileList.length === 0">
            上传并识别
        </el-button>
    </div>
</template>

<script>
import { checkFileFormat, formatFileSize } from '@/utils/fileUtils'
import { uploadAndRecognizeFile } from '@/api/file'

export default {
  name: 'FileUploader',
  props: {
    knowledgeId: {
      type: [Number, String],
      default: ''
    }
  },
  data() {
    return {
      fileList: []
    }
  },
  methods: {
    formatFileSize,
    beforeFileUpload(file) {
      if (!checkFileFormat(file)) {
        this.$message.warning(`文件${file.name}格式不支持，请选择PDF/Word/Excel/TXT/Markdown`)
        return false
      }
      if (file.size > 20 * 1024 * 1024) {
        this.$message.warning(`文件${file.name}超过20MB，无法上传`)
        return false
      }
      return true
    },
    handleFileSelect(file) {
      const isDuplicate = this.fileList.some(item => item.name === file.name && item.size === file.size)
      if (!isDuplicate) {
        this.fileList.push(file)
      }
    },
    removeFile(index) {
      this.fileList.splice(index, 1)
    },
    async handleUpload() {
      if (!this.knowledgeId) {
        this.$message.warning('请先选择知识库')
        return
      }

      const formData = new FormData()
      // Element-UI 的 file 对象中，原始浏览器 File 在 file.raw 上
      this.fileList.forEach(file => formData.append('files', file.raw || file))
      formData.append('knowledgeId', this.knowledgeId)

      const loading = this.$loading({
        lock: true,
        text: '文件识别中...',
        spinner: 'el-icon-loading'
      })

      try {
        await uploadAndRecognizeFile(formData)
        this.$emit('upload-success', this.fileList.map(file => ({
          name: file.name,
          size: this.formatFileSize(file.size)
        })))
        this.fileList = []
      } catch (error) {
        this.$message.error('文件识别失败，请重试')
      } finally {
        loading.close()
      }
    }
  }
}
</script>

<style scoped>
    .file-uploader {
        width: 100%;
    }

    .upload-container {
        margin-bottom: 10px;
    }

    .file-list {
        max-height: 150px;
        overflow-y: auto;
        margin-bottom: 10px;
        border: 1px dashed #eee;
        border-radius: 4px;
        padding: 10px;
    }

    .file-item {
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 5px 0;
        border-bottom: 1px solid #f5f5f5;
    }

    .file-name {
        flex: 1;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
        font-size: 13px;
    }

    .file-size {
        font-size: 12px;
        color: #999;
        margin: 0 10px;
    }

    .delete-btn {
        color: #f56c6c;
    }

    .upload-btn {
        width: 100%;
    }
</style>