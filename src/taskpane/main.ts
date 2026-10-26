import ui from '@nuxt/ui/vue-plugin'
import { createApp } from 'vue'
import App from './App.vue'
import './style.css'

// Office.js APIs are only usable once the host has initialized the add-in.
void Office.onReady(() => {
  createApp(App).use(ui).mount('#app')
})
