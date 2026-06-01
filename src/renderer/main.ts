import { createApp } from 'vue'
import SimpleBar from 'simplebar'
import App from './App.vue'
import 'remixicon/fonts/remixicon.css'
import 'simplebar/dist/simplebar.css'

type SimpleBarElement = HTMLElement & {
  _simplebar?: SimpleBar
}

const app = createApp(App)

app.directive('simplebar', {
  mounted(element: SimpleBarElement) {
    element._simplebar = new SimpleBar(element, {
      autoHide: false,
      scrollbarMinSize: 36
    })
  },
  updated(element: SimpleBarElement) {
    element._simplebar?.recalculate()
  },
  unmounted(element: SimpleBarElement) {
    element._simplebar?.unMount()
    delete element._simplebar
  }
})

app.mount('#app')
