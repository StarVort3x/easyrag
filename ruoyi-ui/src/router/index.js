import Vue from 'vue'
import Router from 'vue-router'
import RagChat from '@/views/rag/chat/index'
import RagKnowledge from '@/views/rag/knowledge/index'

Vue.use(Router)

export default new Router({
    routes: [
        {
            path: '/',
            redirect: '/rag/chat' // 默认跳转到聊天页
        },
        {
            path: '/rag/chat',
            name: 'RagChat',
            component: RagChat,
            meta: { title: 'RAG智能助手' }
        },
        {
            path: '/rag/knowledge',
            name: 'RagKnowledge',
            component: RagKnowledge,
            meta: { title: '知识库管理' }
        }

    ]
})