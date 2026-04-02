import { createRouter, createWebHistory } from 'vue-router'
import MaterialMergeView from '../views/MaterialMergeView.vue'

const router = createRouter({
  history: createWebHistory(import.meta.env.BASE_URL),
  routes: [
    {
      path: '/',
      name: 'home',
      component: MaterialMergeView,
    },
  ],
})

export default router
