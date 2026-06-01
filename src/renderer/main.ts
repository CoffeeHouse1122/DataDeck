import { createApp } from 'vue'
import SimpleBar from 'simplebar'
import App from './App.vue'
import 'remixicon/fonts/remixicon.css'
import 'simplebar/dist/simplebar.css'

type SimpleBarElement = HTMLElement & {
  _simplebar?: SimpleBar
  _simplebarFrame?: number
  _simplebarMutationObserver?: MutationObserver
  _simplebarResizeObserver?: ResizeObserver
}

const app = createApp(App)

function scheduleSimpleBarRecalculate(element: SimpleBarElement): void {
  if (element._simplebarFrame) {
    window.cancelAnimationFrame(element._simplebarFrame)
  }

  element._simplebarFrame = window.requestAnimationFrame(() => {
    element._simplebar?.recalculate()
    element._simplebarFrame = undefined
  })
}

app.directive('simplebar', {
  mounted(element: SimpleBarElement) {
    element._simplebar = new SimpleBar(element, {
      autoHide: false,
      scrollbarMinSize: 36
    })
    element._simplebarResizeObserver = new ResizeObserver(() => scheduleSimpleBarRecalculate(element))
    element._simplebarResizeObserver.observe(element)
    element._simplebarMutationObserver = new MutationObserver(() => scheduleSimpleBarRecalculate(element))
    element._simplebarMutationObserver.observe(element, {
      childList: true,
      subtree: true,
      characterData: true
    })
    scheduleSimpleBarRecalculate(element)
  },
  updated(element: SimpleBarElement) {
    scheduleSimpleBarRecalculate(element)
  },
  unmounted(element: SimpleBarElement) {
    if (element._simplebarFrame) {
      window.cancelAnimationFrame(element._simplebarFrame)
    }
    element._simplebarResizeObserver?.disconnect()
    element._simplebarMutationObserver?.disconnect()
    element._simplebar?.unMount()
    delete element._simplebarFrame
    delete element._simplebarResizeObserver
    delete element._simplebarMutationObserver
    delete element._simplebar
  }
})

app.mount('#app')
