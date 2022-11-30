import Vue from 'vue'
import VueRouter from 'vue-router'
import HelloWorld from '../components/HelloWorld.vue'
import LuckysheetDemo from '../components/LuckysheetDemo.vue'

Vue.use(VueRouter)

const routes = [
  { path: '/', name: 'HelloWorld', component: HelloWorld },
  { path: '/demo', name: 'LuckysheetDemo', component: LuckysheetDemo }
]

const router = new VueRouter({
  routes
})
export default router
