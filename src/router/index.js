import Vue from 'vue'
import VueRouter from 'vue-router'
// import HomeView from '../views/HomeView.vue'

Vue.use(VueRouter)

const routes = [
  {
    path: '/',
    redirect: '/bmi-calculator' // set default page
  },
  {
    path: '/bmi-calculator',
    name: 'bmi-calculator',
    component: () => import('../views/BmiCalculator.vue')
  },
]

const router = new VueRouter({
  mode: 'history',
  base: process.env.BASE_URL,
  routes
})

export default router
