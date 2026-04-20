<template>
    <div class="rag-chat-page">
        <!-- 顶部栏 -->
        <el-header class="page-header">
            <div class="header-left">
                <el-page-header @back="goBack" content="RAG智能助手"></el-page-header>
            </div>
            <div class="header-right">
                <el-select v-model="selectedKnowledgeId"
                           placeholder="选择知识库"
                           size="mini"
                           style="min-width: 180px">
                    <el-option v-for="kb in knowledgeOptions"
                               :key="kb.id"
                               :label="kb.name"
                               :value="kb.id" />
                </el-select>
                <el-button type="default" icon="el-icon-collection" size="mini" @click="goKnowledge">
                    知识库管理
                </el-button>
                <el-dropdown @command="handleToolCommand">
                    <el-button type="default" icon="el-icon-setting" size="mini">工具</el-button>
                    <el-dropdown-menu slot="dropdown">
                        <el-dropdown-item command="history">对话历史</el-dropdown-item>
                        <el-dropdown-item command="clear">清空对话</el-dropdown-item>
                    </el-dropdown-menu>
                </el-dropdown>
                <el-button type="primary" icon="el-icon-refresh" size="mini" @click="refreshPage">刷新</el-button>
            </div>
        </el-header>

        <!-- 新建对话对话框 -->
        <el-dialog title="新建对话" :visible.sync="showNewChatDialog" width="400px">
            <el-form :model="newChatForm" label-width="80px">
                <el-form-item label="对话名称">
                    <el-input v-model="newChatForm.title" placeholder="请输入对话名称"></el-input>
                </el-form-item>
                <el-form-item label="选择知识库">
                    <el-select v-model="newChatForm.knowledgeId" placeholder="请选择知识库">
                        <el-option v-for="kb in knowledgeOptions"
                                   :key="kb.id"
                                   :label="kb.name"
                                   :value="kb.id" />
                    </el-select>
                </el-form-item>
            </el-form>
            <span slot="footer" class="dialog-footer">
                <el-button @click="showNewChatDialog = false">取消</el-button>
                <el-button type="primary" @click="confirmCreateChat">创建</el-button>
            </span>
        </el-dialog>

        <!-- 主体区域 -->
        <el-container class="page-body">
            <!-- 左侧：对话列表 -->
            <el-aside class="page-aside" width="250px">
                <div class="aside-header">
                    <h3>对话列表</h3>
                    <el-button type="primary" icon="el-icon-plus" size="mini" @click="openNewChatDialog">
                        新建对话
                    </el-button>
                </div>
                <div class="chat-list-container">
                    <div v-for="chat in chatList" 
                         :key="chat.id"
                         :class="['chat-item', { active: chat.id === currentChatId }]"
                         @click="switchChat(chat.id)">
                        <div class="chat-title">{{ chat.title }}</div>
                        <div class="chat-kb">知识库: {{ getKnowledgeNameById(chat.knowledgeId) }}</div>
                        <el-button type="text" size="mini" icon="el-icon-delete" @click.stop="deleteChat(chat.id)"></el-button>
                    </div>
                </div>
            </el-aside>

            <!-- 右侧：聊天区域 -->
            <el-main class="page-main">
                <div class="chat-container">
                    <!-- 对话历史 -->
                    <ChatHistory :chat-list="chatHistory" />

                    <!-- 输入区域 -->
                    <div class="input-area">
                        <el-input v-model="inputContent"
                                  type="textarea"
                                  :rows="4"
                                  placeholder="请输入问题..."
                                  @keyup.enter.native="sendMessage"></el-input>

                        <div class="input-actions">
                            <el-button type="default" icon="el-icon-picture-outline" size="mini" @click="openImageUpload">
                                图片
                            </el-button>
                            <el-button type="primary"
                                       icon="el-icon-send"
                                       size="mini"
                                       @click="sendMessage"
                                       :disabled="!inputContent.trim()">

                                发送
                            </el-button>
                        </div>
                    </div>
                </div>
            </el-main>
        </el-container>
    </div>
</template>

<script>
import ChatHistory from '@/components/chat/ChatHistory.vue'
import { sendChatMessage, getChatHistory, clearChatHistory, createChat, deleteChat as deleteChatAPI, getChatList, getChatMessages } from '@/api/chat'
import { getKnowledgeList } from '@/api/knowledge'

export default {
  name: 'RagChat',
  components: { ChatHistory },
  data() {
    return {
      // 聊天数据
      chatHistory: [],
      inputContent: '',
      knowledgeOptions: [],
      selectedKnowledgeId: null,
      // 多对话数据
      chatList: [],
      currentChatId: null,
      chatCounter: 0,
      // 新建对话对话框
      showNewChatDialog: false,
      newChatForm: {
        title: '',
        knowledgeId: null
      }
    }
  },
  async mounted() {
    await this.loadKnowledgeOptions()
    await this.loadChatList()
  },
  methods: {

    // 返回上一页
    goBack() {
      this.$router.go(-1)
    },

    // 跳转到知识库管理页面
    goKnowledge() {
      this.$router.push('/rag/knowledge')
    },

    // 刷新页面
    refreshPage() {
      window.location.reload()
    },

    // 工具菜单命令
    async handleToolCommand(cmd) {
      if (cmd === 'history') {
        try {
          const res = await getChatHistory()
          this.chatHistory = res.data.map(msg => ({
            role: msg.role === 'assistant' ? 'ai' : msg.role,
            content: msg.content
          }))
          this.$message.success('已加载对话历史')
        } catch (error) {
          console.error('加载对话历史失败:', error)
          this.$message.error('加载对话历史失败')
        }
      } else if (cmd === 'clear') {
        try {
          await clearChatHistory()
          this.chatHistory = []
          this.$message.success('对话已清空')
        } catch (error) {
          console.error('清空失败:', error)
          this.$message.error('清空失败，请重试')
        }
      }
    },

    // 加载知识库选项
    async loadKnowledgeOptions() {
      try {
        const res = await getKnowledgeList()
        const data = Array.isArray(res) ? res : (res.data || res)
        
        // 展平嵌套的知识库结构，为子知识库添加缩进前缀
        this.knowledgeOptions = []
        data.forEach(kb => {
          // 添加主知识库
          this.knowledgeOptions.push({
            id: kb.id,
            name: kb.label || kb.name,
            label: kb.label || kb.name,
            isParent: true
          })
          // 添加子知识库，使用 "  └─ " 前缀表示层级
          if (kb.children && Array.isArray(kb.children)) {
            kb.children.forEach(child => {
              this.knowledgeOptions.push({
                id: child.id,
                name: `  └─ ${child.label || child.name}`,
                label: `  └─ ${child.label || child.name}`,
                isChild: true
              })
            })
          }
        })
        
        if (this.knowledgeOptions.length && !this.selectedKnowledgeId) {
          this.selectedKnowledgeId = this.knowledgeOptions[0].id
        }
      } catch (error) {
        console.error('加载知识库列表失败:', error)
        this.$message.error('加载知识库列表失败')
      }
    },

    // 发送聊天消息
    async sendMessage() {
      const content = this.inputContent.trim()
      if (!content) return

      if (!this.selectedKnowledgeId) {
        this.$message.warning('请先在右上角选择一个知识库')
        return
      }

      // 添加用户消息
      this.chatHistory.push({ role: 'user', content })
      this.inputContent = ''

      // 显示加载
      const loading = this.$loading({
        lock: false,
        text: 'AI思考中...',
        background: 'rgba(255, 255, 255, 0.5)'
      })

      try {
        const res = await sendChatMessage({
          knowledgeId: this.selectedKnowledgeId,
          content,
          chatId: this.currentChatId  // 传递 chat_id
        })
        // 处理响应数据
        const data = res.data || res
        const aiMessage = {
          role: 'ai',
          content: data.reply || data.content || '无法获取回复',
          thinking: data.thinking || '',
          references: data.references || []
        }
        this.chatHistory.push(aiMessage)
      } catch (error) {
        console.error('获取回复失败:', error)
        this.$message.error('获取回复失败，请重试')
        // 即使失败也添加错误提示消息
        this.chatHistory.push({
          role: 'ai',
          content: '抱歉，处理您的请求时出现错误，请稍后重试。',
          thinking: '处理出错',
          references: []
        })
      } finally {
        loading.close()
      }
    },

    // 打开图片上传（预留功能）
    openImageUpload() {
      this.$message.info('图片上传功能待实现')
    },

    // 打开新建对话对话框
    openNewChatDialog() {
      this.newChatForm = {
        title: '',
        knowledgeId: this.selectedKnowledgeId || (this.knowledgeOptions.length > 0 ? this.knowledgeOptions[0].id : null)
      }
      this.showNewChatDialog = true
    },

    // 确认创建新对话
    async confirmCreateChat() {
      if (!this.newChatForm.title.trim()) {
        this.$message.warning('请输入对话名称')
        return
      }
      if (!this.newChatForm.knowledgeId) {
        this.$message.warning('请选择知识库')
        return
      }

      try {
        // 调用后端 API 创建对话
        const res = await createChat({
          title: this.newChatForm.title.trim(),
          knowledgeId: this.newChatForm.knowledgeId
        })

        const chatData = res.data || res
        const newChat = {
          id: chatData.chatId,
          title: chatData.title,
          knowledgeId: chatData.knowledgeId,
          messages: [],
          createdAt: new Date()
        }
        this.chatList.push(newChat)
        this.switchChat(newChat.id)
        this.showNewChatDialog = false
        this.$message.success('新对话已创建')
      } catch (error) {
        console.error('创建对话失败:', error)
        this.$message.error('创建对话失败，请重试')
      }
    },

    // 切换对话
    async switchChat(chatId) {
      const chat = this.chatList.find(c => c.id === chatId)
      if (!chat) return

      // 保存当前对话的消息
      if (this.currentChatId) {
        const currentChat = this.chatList.find(c => c.id === this.currentChatId)
        if (currentChat) {
          currentChat.messages = this.chatHistory
          currentChat.knowledgeId = this.selectedKnowledgeId
        }
      }

      // 切换到新对话
      this.currentChatId = chatId
      this.selectedKnowledgeId = chat.knowledgeId
      this.inputContent = ''

      // 从后端加载对话消息
      try {
        const res = await getChatMessages(chatId)
        const data = res.data || res
        this.chatHistory = Array.isArray(data) ? data.map(msg => ({
          role: msg.role === 'assistant' ? 'ai' : msg.role,
          content: msg.content,
          thinking: msg.thinking || '',
          references: msg.references || []
        })) : []
      } catch (error) {
        console.error('加载对话消息失败:', error)
        this.chatHistory = []
      }
    },

    // 删除对话
    deleteChat(chatId) {
      this.$confirm('确定要删除这个对话吗?', '提示', {
        confirmButtonText: '确定',
        cancelButtonText: '取消',
        type: 'warning'
      }).then(() => {
        const index = this.chatList.findIndex(c => c.id === chatId)
        if (index > -1) {
          this.chatList.splice(index, 1)
          
          // 如果删除的是当前对话，切换到第一个对话
          if (this.currentChatId === chatId) {
            if (this.chatList.length > 0) {
              this.switchChat(this.chatList[0].id)
            } else {
              this.currentChatId = null
              this.chatHistory = []
              this.inputContent = ''
            }
          }
          this.$message.success('对话已删除')
        }
      }).catch(() => {})
    },

    // 加载对话列表
    async loadChatList() {
      try {
        const res = await getChatList()
        const data = res.data || res
        this.chatList = Array.isArray(data) ? data : []
        
        // 如果有对话，自动选择第一个
        if (this.chatList.length > 0) {
          this.switchChat(this.chatList[0].id)
        }
      } catch (error) {
        console.error('加载对话列表失败:', error)
        // 不显示错误提示，因为可能是第一次使用
      }
    },

    // 根据知识库ID获取知识库名称
    getKnowledgeNameById(kbId) {
      const kb = this.knowledgeOptions.find(k => k.id === kbId)
      return kb ? kb.name : '未知'
    }
  }
}
</script>

<style scoped>
    .rag-chat-page {
        height: 100vh;
        box-sizing: border-box;
        overflow: hidden;
    }

    /* 顶部栏 */
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
        gap: 10px;
    }

    /* 主体区域 */
    .page-body {
        height: calc(100vh - 60px);
    }

    .page-aside {
        border-right: 1px solid #eee;
        padding: 15px;
        box-sizing: border-box;
    }

    .aside-card {
        margin-bottom: 15px;
    }

    .card-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
    }

    /* 聊天区域 */
    .page-main {
        padding: 15px;
        box-sizing: border-box;
        overflow: hidden;
    }

    .chat-container {
        height: 100%;
        display: flex;
        flex-direction: column;
    }

    .input-area {
        padding-top: 15px;
        border-top: 1px solid #eee;
    }

    .input-actions {
        display: flex;
        justify-content: flex-end;
        gap: 10px;
        margin-top: 10px;
    }

    /* 左侧对话列表样式 */
    .aside-header {
        display: flex;
        justify-content: space-between;
        align-items: center;
        margin-bottom: 15px;
        padding-bottom: 10px;
        border-bottom: 1px solid #eee;
    }

    .aside-header h3 {
        margin: 0;
        font-size: 14px;
        font-weight: 600;
    }

    .chat-list-container {
        height: calc(100vh - 200px);
        overflow-y: auto;
        display: flex;
        flex-direction: column;
        gap: 8px;
    }

    .chat-item {
        padding: 10px;
        border-radius: 6px;
        cursor: pointer;
        background-color: #f5f5f5;
        transition: all 0.3s ease;
        position: relative;
    }

    .chat-item:hover {
        background-color: #efefef;
    }

    .chat-item.active {
        background-color: #409eff;
        color: white;
    }

    .chat-title {
        font-weight: 600;
        font-size: 13px;
        margin-bottom: 4px;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
    }

    .chat-kb {
        font-size: 12px;
        color: #999;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
    }

    .chat-item.active .chat-kb {
        color: rgba(255, 255, 255, 0.8);
    }

    .chat-item .el-button {
        position: absolute;
        right: 5px;
        top: 5px;
        opacity: 0;
        transition: opacity 0.3s ease;
    }

    .chat-item:hover .el-button {
        opacity: 1;
    }
</style>