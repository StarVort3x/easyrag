<template>
  <div class="rag-knowledge-page">
    <el-header class="page-header">
      <div class="header-left">
        <el-page-header @back="$router.go(-1)" content="知识库管理"></el-page-header>
      </div>
      <div class="header-right">
        <el-select v-model="currentKnowledgeId"
                   placeholder="选择知识库"
                   size="mini"
                   style="min-width: 200px"
                   @change="handleKnowledgeChange">
          <el-option v-for="kb in knowledgeList"
                     :key="kb.id"
                     :label="kb.name"
                     :value="kb.id" />
        </el-select>
        <el-button type="danger" size="mini" icon="el-icon-delete" :disabled="!currentKnowledgeId" @click="deleteCurrentKnowledge">删除知识库</el-button>
        <el-button type="primary" size="mini" icon="el-icon-plus" @click="createKnowledge">新建知识库</el-button>
      </div>
    </el-header>
    <el-main class="page-main">
      <el-row :gutter="20">
        <el-col :span="16">
          <el-card shadow="hover">
            <div slot="header" class="card-header">
              <span>当前知识库文件</span>
              <el-button type="danger"
                         size="mini"
                         :disabled="multipleSelection.length === 0"
                         @click="handleBatchDelete">批量删除</el-button>
            </div>
            <el-table :data="fileList"
                      border
                      size="small"
                      @selection-change="handleSelectionChange">
              <el-table-column type="selection" width="50" />
              <el-table-column prop="filename" label="文件名" min-width="260" show-overflow-tooltip />
              <el-table-column prop="size" label="大小" width="120">
                <template slot-scope="scope">
                  {{ formatFileSize(scope.row.size) }}
                </template>
              </el-table-column>
              <el-table-column prop="create_time" label="上传时间" width="180" />
              <el-table-column label="操作" width="100">
                <template slot-scope="scope">
                  <el-button type="text" size="mini" @click="deleteSingle(scope.row)">删除</el-button>
                </template>
              </el-table-column>
            </el-table>
          </el-card>
        </el-col>
        <el-col :span="8">
          <el-card shadow="hover">
            <div slot="header" class="card-header">
              <span>上传文档构建知识库</span>
            </div>
            <p class="tip">选择知识库后上传的文档会写入对应知识库，并参与该知识库的 RAG 检索。</p>
            <FileUploader :knowledge-id="currentKnowledgeId" @upload-success="handleUploadSuccess" />
          </el-card>
        </el-col>
      </el-row>
    </el-main>
  </div>
  </template>

<script>
import FileUploader from '@/components/file/FileUploader.vue'
import { getKnowledgeList, addKnowledge, deleteKnowledge } from '@/api/knowledge'
import { getFileList, deleteFile } from '@/api/file'
import { formatFileSize } from '@/utils/fileUtils'

export default {
  name: 'RagKnowledge',
  components: { FileUploader },
  data() {
    return {
      knowledgeList: [],
      currentKnowledgeId: null,
      fileList: [],
      multipleSelection: []
    }
  },
  created() {
    this.loadKnowledgeList()
  },
  methods: {
    formatFileSize,

    async loadKnowledgeList() {
      try {
        const res = await getKnowledgeList()
        const data = Array.isArray(res) ? res : (res.data || res)
        
        // 展平嵌套的知识库结构，为子知识库添加缩进前缀
        this.knowledgeList = []
        data.forEach(kb => {
          // 添加主知识库
          this.knowledgeList.push({
            id: kb.id,
            name: kb.label || kb.name,
            label: kb.label || kb.name,
            isParent: true
          })
          // 添加子知识库，使用 "  └─ " 前缀表示层级
          if (kb.children && Array.isArray(kb.children)) {
            kb.children.forEach(child => {
              this.knowledgeList.push({
                id: child.id,
                name: `  └─ ${child.label || child.name}`,
                label: `  └─ ${child.label || child.name}`,
                isChild: true
              })
            })
          }
        })
        
        if (this.knowledgeList.length && !this.currentKnowledgeId) {
          this.currentKnowledgeId = this.knowledgeList[0].id
          this.loadFileList()
        }
      } catch (e) {
        console.error('加载知识库失败', e)
        this.$message.error('加载知识库失败')
      }
    },

    async loadFileList() {
      if (!this.currentKnowledgeId) return
      try {
        const res = await getFileList({ knowledgeId: this.currentKnowledgeId })
        this.fileList = Array.isArray(res) ? res : (res.data || res)
      } catch (e) {
        console.error('加载文件列表失败', e)
        this.$message.error('加载文件列表失败')
      }
    },

    handleKnowledgeChange() {
      this.loadFileList()
    },

    async createKnowledge() {
      this.$prompt('请输入知识库名称', '新建知识库', {
        confirmButtonText: '确定',
        cancelButtonText: '取消',
        inputPattern: /.+/,
        inputErrorMessage: '名称不能为空'
      }).then(async ({ value }) => {
        try {
          const res = await addKnowledge({ name: value })
          this.$message.success('创建成功')
          await this.loadKnowledgeList()
          if (res && res.id) {
            this.currentKnowledgeId = res.id
            this.loadFileList()
          }
        } catch (e) {
          console.error('创建知识库失败', e)
          this.$message.error('创建知识库失败')
        }
      }).catch(() => {})
    },

    async deleteCurrentKnowledge() {
      if (!this.currentKnowledgeId) {
        this.$message.warning('请先选择要删除的知识库')
        return
      }

      const currentKb = this.knowledgeList.find(kb => kb.id === this.currentKnowledgeId)
      const kbName = currentKb ? currentKb.name : '知识库'

      try {
        await this.$confirm(`确定删除知识库「${kbName}」吗？删除后无法恢复。`, '删除知识库', {
          confirmButtonText: '确定删除',
          cancelButtonText: '取消',
          type: 'warning'
        })

        await deleteKnowledge(this.currentKnowledgeId)
        this.$message.success('知识库已删除')
        
        // 重新加载知识库列表
        await this.loadKnowledgeList()
        
        // 如果还有知识库，选择第一个；否则清空
        if (this.knowledgeList.length > 0) {
          this.currentKnowledgeId = this.knowledgeList[0].id
          this.loadFileList()
        } else {
          this.currentKnowledgeId = null
          this.fileList = []
        }
      } catch (e) {
        if (e !== 'cancel') {
          console.error('删除知识库失败', e)
          this.$message.error('删除知识库失败')
        }
      }
    },

    handleSelectionChange(val) {
      this.multipleSelection = val
    },

    async deleteSingle(row) {
      try {
        await this.$confirm(`确定删除文件「${row.filename}」吗？`, '提示', {
          type: 'warning'
        })
        await deleteFile([row.id])
        this.$message.success('删除成功')
        this.loadFileList()
      } catch (e) {
        if (e !== 'cancel') {
          console.error('删除文件失败', e)
          this.$message.error('删除文件失败')
        }
      }
    },

    async handleBatchDelete() {
      if (!this.multipleSelection.length) return
      const ids = this.multipleSelection.map(item => item.id)
      try {
        await this.$confirm(`确定删除选中的 ${ids.length} 个文件吗？`, '提示', {
          type: 'warning'
        })
        await deleteFile(ids)
        this.$message.success('删除成功')
        this.loadFileList()
      } catch (e) {
        if (e !== 'cancel') {
          console.error('批量删除失败', e)
          this.$message.error('批量删除失败')
        }
      }
    },

    handleUploadSuccess(files) {
      this.$message.success(`已成功上传并识别 ${files.length} 个文件`)
      this.loadFileList()
    }
  }
}
</script>

<style scoped>
.rag-knowledge-page {
  height: 100vh;
  box-sizing: border-box;
  overflow: hidden;
}

.page-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  padding: 0 20px;
  border-bottom: 1px solid #eee;
  height: 60px;
}

.header-right {
  display: flex;
  align-items: center;
  gap: 10px;
}

.page-main {
  padding: 20px;
  box-sizing: border-box;
}

.card-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
}

.tip {
  margin-bottom: 16px;
  color: #909399;
  font-size: 13px;
}
</style>
