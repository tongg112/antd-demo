import Vue from 'vue'
import App from './App.vue'
import Antd from 'ant-design-vue'
import 'ant-design-vue/dist/antd.css'

// 图片预览 注意版本 1.5.1
import 'viewerjs/dist/viewer.css'
import Viewer from 'v-viewer'

Vue.config.productionTip = false

Vue.use(Antd)

// 引入图片预览
Vue.use(Viewer, { name: 'v-viewer' })

new Vue({
  render: h => h(App),
}).$mount('#app')
