<template>
    <div :class="`message-item ${role === 'user' ? 'user-message' : 'ai-message'}`">
        <div class="avatar">
            {{ role === 'user' ? '我' : 'AI' }}
        </div>
        <div class="message-content">
            <!-- 主要内容 - 支持 Markdown 格式 -->
            <div class="content-text markdown-body" v-html="formatContent(content)"></div>
            
            <!-- AI回复的思考过程和引用文献 -->
            <div v-if="role === 'ai' && (thinking || references.length > 0)" class="ai-details">
                <!-- 思考过程 - 可折叠 -->
                <div v-if="thinking" class="collapsible-section">
                    <div class="section-header" @click="toggleThinking">
                        <i :class="showThinking ? 'el-icon-arrow-down' : 'el-icon-arrow-right'"></i>
                        <i class="el-icon-info"></i>
                        <span>思考过程</span>
                    </div>
                    <transition name="collapse">
                        <div v-show="showThinking" class="thinking-content">{{ thinking }}</div>
                    </transition>
                </div>
                
                <!-- 引用文献 - 可折叠 -->
                <div v-if="references.length > 0" class="collapsible-section">
                    <div class="section-header" @click="toggleReferences">
                        <i :class="showReferences ? 'el-icon-arrow-down' : 'el-icon-arrow-right'"></i>
                        <i class="el-icon-document"></i>
                        <span>引用文献 ({{ references.length }})</span>
                    </div>
                    <transition name="collapse">
                        <div v-show="showReferences" class="references-list">
                            <div v-for="(ref, idx) in references" :key="idx" class="reference-item">
                                <span class="ref-number">[{{ idx + 1 }}]</span>
                                <span class="ref-content">{{ ref }}</span>
                            </div>
                        </div>
                    </transition>
                </div>
            </div>
            
            <!-- 文件识别结果附加信息 -->
            <div class="file-attach" v-if="type === 'file'">
                <el-tag type="info" size="mini">文件识别完成</el-tag>
                <span class="file-name">{{ fileName }}</span>
            </div>
        </div>
    </div>
</template>

<script>
export default {
  name: 'MessageItem',
  props: {
    role: { // 角色：user/ai
      type: String,
      required: true,
      validator: val => ['user', 'ai'].includes(val)
    },
    content: { // 消息内容
      type: String,
      required: true
    },
    type: { // 消息类型：text/file
      type: String,
      default: 'text'
    },
    fileName: { // 文件名（仅type=file时生效）
      type: String,
      default: ''
    },
    thinking: { // 思考过程（AI回复时）
      type: String,
      default: ''
    },
    references: { // 引用文献列表（AI回复时）
      type: Array,
      default: () => []
    }
  },
  data() {
    return {
      showThinking: false,
      showReferences: false
    }
  },
  methods: {
    // 格式化内容为 Markdown HTML
    formatContent(text) {
      if (!text) return ''
      
      let html = text
      
      // 转义 HTML 特殊字符
      html = html
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;')
      
      // 处理代码块 ```...```
      html = html.replace(/```([\s\S]*?)```/g, '<pre><code>$1</code></pre>')
      
      // 处理行内代码 `...`
      html = html.replace(/`([^`]+)`/g, '<code class="inline-code">$1</code>')
      
      // 处理加粗 **...**
      html = html.replace(/\*\*([^\*]+)\*\*/g, '<strong>$1</strong>')
      
      // 处理斜体 *...*
      html = html.replace(/\*([^\*]+)\*/g, '<em>$1</em>')
      
      // 处理标题 # ... ## ... ### ...
      html = html.replace(/^### (.*?)$/gm, '<h3>$1</h3>')
      html = html.replace(/^## (.*?)$/gm, '<h2>$1</h2>')
      html = html.replace(/^# (.*?)$/gm, '<h1>$1</h1>')
      
      // 处理列表 - ...
      html = html.replace(/^- (.*?)$/gm, '<li>$1</li>')
      html = html.replace(/(<li>.*<\/li>)/s, '<ul>$1</ul>')
      
      // 处理换行
      html = html.replace(/\n/g, '<br>')
      
      return html
    },
    
    // 切换思考过程显示
    toggleThinking() {
      this.showThinking = !this.showThinking
    },
    
    // 切换引用文献显示
    toggleReferences() {
      this.showReferences = !this.showReferences
    }
  }
}
</script>

<style scoped>
    .message-item {
        display: flex;
        align-items: flex-start;
        margin-bottom: 15px;
        max-width: 80%;
    }

    .user-message {
        flex-direction: row-reverse;
        margin-left: auto;
    }

    .avatar {
        width: 36px;
        height: 36px;
        border-radius: 50%;
        color: #fff;
        line-height: 36px;
        text-align: center;
        font-size: 14px;
    }

    .user-message .avatar {
        background-color: #409eff;
        margin-left: 10px;
    }

    .ai-message .avatar {
        background-color: #67c23a;
        margin-right: 10px;
    }

    .message-content {
        padding: 10px 14px;
        border-radius: 8px;
        line-height: 1.6;
    }

    .user-message .message-content {
        background-color: #409eff;
        color: #fff;
    }

    .ai-message .message-content {
        background-color: #f5f5f5;
        color: #333;
    }

    .file-attach {
        margin-top: 8px;
        padding-top: 8px;
        border-top: 1px dashed #eee;
        display: flex;
        align-items: center;
        gap: 8px;
        font-size: 13px;
    }

    .file-name {
        color: #666;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
        max-width: 200px;
    }

    /* AI详细信息样式 */
    .ai-details {
        margin-top: 12px;
        padding-top: 12px;
        border-top: 1px solid rgba(0, 0, 0, 0.1);
    }

    .thinking-section,
    .references-section {
        margin-bottom: 10px;
    }

    .section-title {
        font-size: 12px;
        font-weight: 600;
        color: #666;
        margin-bottom: 6px;
        display: flex;
        align-items: center;
        gap: 4px;
    }

    .section-title i {
        font-size: 12px;
    }

    .thinking-content {
        font-size: 12px;
        color: #666;
        line-height: 1.6;
        padding: 8px 10px;
        background: linear-gradient(135deg, rgba(64, 158, 255, 0.05) 0%, rgba(103, 194, 58, 0.05) 100%);
        border-left: 3px solid #409eff;
        border-radius: 4px;
        word-break: break-word;
    }

    .references-list {
        font-size: 12px;
        color: #666;
    }

    .reference-item {
        display: flex;
        gap: 6px;
        margin-bottom: 4px;
        line-height: 1.4;
    }

    .ref-number {
        color: #409eff;
        font-weight: 600;
        flex-shrink: 0;
    }

    .ref-content {
        word-break: break-word;
        color: #888;
    }

    /* Markdown 样式 */
    .markdown-body {
        word-break: break-word;
        white-space: pre-wrap;
        line-height: 1.8;
    }

    .markdown-body h1,
    .markdown-body h2,
    .markdown-body h3 {
        margin: 12px 0 8px 0;
        font-weight: 600;
    }

    .markdown-body h1 {
        font-size: 18px;
        border-bottom: 2px solid #409eff;
        padding-bottom: 6px;
    }

    .markdown-body h2 {
        font-size: 16px;
        border-bottom: 1px solid #eee;
        padding-bottom: 4px;
    }

    .markdown-body h3 {
        font-size: 14px;
    }

    .markdown-body strong {
        font-weight: 700;
        color: #333;
    }

    .markdown-body em {
        font-style: italic;
        color: #666;
    }

    .markdown-body code {
        background-color: rgba(0, 0, 0, 0.05);
        padding: 2px 6px;
        border-radius: 3px;
        font-family: 'Monaco', 'Menlo', 'Ubuntu Mono', monospace;
        font-size: 13px;
    }

    .markdown-body code.inline-code {
        color: #e83e8c;
    }

    .markdown-body pre {
        background-color: rgba(0, 0, 0, 0.08);
        padding: 10px 12px;
        border-radius: 4px;
        overflow-x: auto;
        margin: 8px 0;
    }

    .markdown-body pre code {
        background-color: transparent;
        padding: 0;
        color: #333;
    }

    .markdown-body ul {
        margin: 8px 0;
        padding-left: 24px;
    }

    .markdown-body li {
        margin: 4px 0;
        list-style-type: disc;
    }

    .markdown-body br {
        display: block;
        height: 0;
        line-height: 0;
    }

    /* 可折叠部分样式 */
    .collapsible-section {
        margin-top: 10px;
        border: 1px solid #e8e8e8;
        border-radius: 4px;
        overflow: hidden;
    }

    .section-header {
        display: flex;
        align-items: center;
        gap: 6px;
        padding: 8px 10px;
        background-color: #f9f9f9;
        cursor: pointer;
        user-select: none;
        transition: background-color 0.2s;
        font-size: 12px;
        font-weight: 600;
        color: #666;
    }

    .section-header:hover {
        background-color: #f0f0f0;
    }

    .section-header i {
        font-size: 12px;
        transition: transform 0.2s;
    }

    /* 折叠动画 */
    .collapse-enter-active,
    .collapse-leave-active {
        transition: all 0.3s ease;
        max-height: 500px;
        overflow: hidden;
    }

    .collapse-enter,
    .collapse-leave-to {
        max-height: 0;
        opacity: 0;
    }

    .thinking-content {
        font-size: 12px;
        color: #666;
        line-height: 1.6;
        padding: 10px;
        background-color: #fafafa;
        border-top: 1px solid #e8e8e8;
        word-break: break-word;
        white-space: pre-wrap;
    }

    .references-list {
        font-size: 12px;
        color: #666;
        padding: 8px 10px;
        background-color: #fafafa;
        border-top: 1px solid #e8e8e8;
    }

    .reference-item {
        display: flex;
        gap: 8px;
        margin-bottom: 6px;
        line-height: 1.5;
    }

    .reference-item:last-child {
        margin-bottom: 0;
    }

    .ref-number {
        color: #409eff;
        font-weight: 600;
        flex-shrink: 0;
        min-width: 30px;
    }

    .ref-content {
        word-break: break-word;
        color: #888;
        flex: 1;
    }
</style>
