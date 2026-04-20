<template>
    <div class="chat-history-container">
        <!-- 空状态提示 -->
        <div class="empty-tip" v-if="chatList.length === 0">
            <i class="el-icon-info empty-icon"></i>
            <p>请选择知识库或上传文件，开始与RAG助手对话</p>
        </div>
        <!-- 对话列表 -->
        <div class="chat-list" v-else>
            <MessageItem v-for="(msg, index) in chatList"
                         :key="index"
                         :role="msg.role"
                         :content="msg.content"
                         :type="msg.type"
                         :file-name="msg.fileName"
                         :thinking="msg.thinking"
                         :references="msg.references" />
        </div>
    </div>
</template>

<script>
import MessageItem from './MessageItem.vue'

export default {
  name: 'ChatHistory',
  components: { MessageItem },
  props: {
    chatList: { // 对话历史列表
      type: Array,
      default: () => []
    }
  }
}
</script>

<style scoped>
    .chat-history-container {
        height: 100%;
        overflow-y: auto;
        padding: 15px;
        box-sizing: border-box;
    }

    .empty-tip {
        height: 100%;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
        color: #999;
        gap: 10px;
    }

    .empty-icon {
        font-size: 32px;
    }

    .chat-list {
        display: flex;
        flex-direction: column;
        gap: 5px;
    }
</style>