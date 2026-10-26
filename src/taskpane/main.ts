import ui from '@nuxt/ui/vue-plugin'
import { createApp } from 'vue'
import { createAppI18n, detectLocale } from '@/i18n'
import App from './App.vue'
import './style.css'

// Office.js APIs, including the display language, are only usable once the host has initialized the add-in.
void Office.onReady(() => {
  createApp(App)
    .use(createAppI18n(detectLocale()))
    .use(ui)
    .mount('#app')
})
