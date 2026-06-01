<script setup lang="ts">
import { nextTick, onBeforeUnmount, onMounted, onUpdated, ref } from 'vue'
import SimpleBar from 'simplebar'

const props = withDefaults(defineProps<{
  autoHide?: boolean
  contentClass?: string
  scrollbarMinSize?: number
}>(), {
  autoHide: true,
  contentClass: '',
  scrollbarMinSize: 36
})

const root = ref<HTMLElement | null>(null)
const content = ref<HTMLElement | null>(null)
const contentWrapper = ref<HTMLElement | null>(null)

let simplebar: SimpleBar | null = null
let frame = 0
let mutationObserver: MutationObserver | null = null
let resizeObserver: ResizeObserver | null = null

function scheduleRecalculate(): void {
  if (frame) {
    window.cancelAnimationFrame(frame)
  }

  frame = window.requestAnimationFrame(() => {
    simplebar?.recalculate()
    frame = 0
  })
}

onMounted(async () => {
  await nextTick()
  if (!root.value || !content.value || !contentWrapper.value) {
    return
  }

  simplebar = new SimpleBar(root.value, {
    autoHide: props.autoHide,
    contentNode: content.value,
    scrollbarMinSize: props.scrollbarMinSize,
    scrollableNode: contentWrapper.value
  })

  resizeObserver = new ResizeObserver(scheduleRecalculate)
  resizeObserver.observe(root.value)
  resizeObserver.observe(content.value)

  mutationObserver = new MutationObserver(scheduleRecalculate)
  mutationObserver.observe(content.value, {
    childList: true,
    subtree: true,
    characterData: true
  })

  scheduleRecalculate()
})

onUpdated(scheduleRecalculate)

onBeforeUnmount(() => {
  if (frame) {
    window.cancelAnimationFrame(frame)
  }
  resizeObserver?.disconnect()
  mutationObserver?.disconnect()
  simplebar?.unMount()
})
</script>

<template>
  <div ref="root" data-simplebar="init">
    <div class="simplebar-wrapper">
      <div class="simplebar-height-auto-observer-wrapper">
        <div class="simplebar-height-auto-observer" />
      </div>
      <div class="simplebar-mask">
        <div class="simplebar-offset">
          <div
            ref="contentWrapper"
            class="simplebar-content-wrapper"
            role="region"
            aria-label="scrollable content"
            tabindex="0"
          >
            <div ref="content" class="simplebar-content" :class="contentClass">
              <slot />
            </div>
          </div>
        </div>
      </div>
      <div class="simplebar-placeholder" />
    </div>
  </div>
</template>

<style scoped>
.simplebar-content-wrapper {
  overflow-x: hidden !important;
}

.simplebar-content {
  padding-right: var(--simplebar-content-gutter, 0) !important;
}

:deep(.simplebar-track.simplebar-horizontal) {
  display: none;
}

.simplebar-content.simplebar-stack {
  display: grid;
  gap: var(--simplebar-content-gap, 0);
  align-content: start;
}
</style>