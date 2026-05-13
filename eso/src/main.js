import { createApp } from 'vue'
import ElementPlus from 'element-plus'
import 'element-plus/dist/index.css'
import { Document, ChatDotRound } from '@element-plus/icons-vue'
import App from './App.vue'
import './style.css'

const app = createApp(App)

app.use(ElementPlus)
app.component('Document', Document)
app.component('ChatDotRound', ChatDotRound)
app.mount('#app')