import Vue from 'vue'
import App from './App.vue'
import TestMammoth from './TestMammoth.vue'
import TestDocxPreview from './TestDocxPreview.vue'

Vue.config.productionTip = false

// 用 hash 切换页面,避免引入 vue-router
//   /              -> App.vue(原编辑器)
//   /#mammoth      -> TestMammoth.vue
//   /#docx-preview -> TestDocxPreview.vue
const map = {
  '#mammoth': TestMammoth,
  '#docx-preview': TestDocxPreview
}
const Root = map[location.hash] || App

new Vue({
  render: h => h(Root),
}).$mount('#app')

window.addEventListener('hashchange', () => location.reload())
