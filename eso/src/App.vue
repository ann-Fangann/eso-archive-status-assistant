<template>
  <div class="app-container">
    <header>
      <h1>ESO数据智能体</h1>
      <!-- <p class="subtitle">上传Excel文件，自动导入数据到MySQL数据库</p> -->
    </header>
    
    <div class="main-layout">
      <aside class="sidebar">
        <el-menu
          :default-active="activeComponent"
          class="sidebar-menu"
          @select="activeComponent = $event"
        >
          <el-menu-item index="excel">
            <el-icon><Document /></el-icon>
            <span>多源Excel数据融合平台</span>
          </el-menu-item>
          <el-menu-item index="chatbot">
            <el-icon><ChatDotRound /></el-icon>
            <span>智能问答</span>
          </el-menu-item>
        </el-menu>
      </aside>
      
      <main class="main-content">
        <div v-show="activeComponent === 'excel'" class="content-wrapper">
          <ExcelMatcher />
        </div>
        <div v-show="activeComponent === 'chatbot'" class="content-wrapper">
          <ChatBot />
        </div>
      </main>
    </div>
    
    
  </div>
</template>

<script setup>
import { ref } from 'vue'
import ExcelMatcher from './components/ExcelMatcher.vue'
import ChatBot from './components/ChatBot.vue'

const activeComponent = ref('excel')
</script>

<style scoped>
html, body {
  height: 100%;
  margin: 0;
  padding: 0;
}

#app {
  height: 100%;
}

.app-container {
  min-height: 100vh;
  height: 100%;
  background: linear-gradient(135deg, #f5f7fa 0%, #e4edf5 100%);
  display: flex;
  flex-direction: column;
  width: 100%;
  padding: 0;
  margin: 0;
  box-sizing: border-box;
}

.main-layout {
  display: flex;
  flex: 1;
  width: 100%;
  overflow: hidden;
}

.sidebar {
  width: 200px;
  background-color: white;
  box-shadow: 2px 0 8px rgba(0,0,0,.1);
  z-index: 10;
  flex-shrink: 0;
  overflow-y: auto;
}

.sidebar-menu {
  border-right: none;
  height: 100%;
}

.main-content {
  flex: 1;
  overflow: hidden;
  width: 100%;
  display: flex;
  flex-direction: column;
}

.content-wrapper {
  flex: 1;
  overflow-y: auto;
  padding: 20px;
  width: 100%;
  box-sizing: border-box;
  height: 100%;
  display: flex;
  flex-direction: column;
}

header {
  background: linear-gradient(90deg, #409eff 0%, #64b5f6 100%);
  color: white;
  padding: 15px 20px;
  text-align: center;
  box-shadow: 0 2px 12px rgba(0,0,0,.1);
  display: flex;
  flex-direction: column;
  justify-content: center;
  align-items: center;
  flex-shrink: 0;
}

header h1 {
  margin: 0;
  font-size: 2rem;
  font-weight: 600;
  text-align: center;
}

.subtitle {
  font-size: 1rem;
  margin-top: 8px;
  opacity: 0.9;
}


@media (max-width: 768px) {
  .main-layout {
    flex-direction: column;
  }
  
  .sidebar {
    width: 100%;
    height: auto;
  }
  
  header h1 {
    font-size: 1.8rem;
  }
  
  .content-wrapper {
    padding: 10px;
  }
  
  .subtitle {
    font-size: 0.9rem;
  }
}
</style>
