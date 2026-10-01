<script setup lang="ts">
import { onBeforeUnmount, onMounted, ref } from 'vue'

defineProps<{ entries: Array<{ label: string; message: string }> }>()
const emit = defineEmits<{ close: [] }>()
const dialog = ref<HTMLDialogElement | null>(null)

function handleKey(event: KeyboardEvent): void {
  if (event.key !== 'Tab') return
  const controls = Array.from(dialog.value?.querySelectorAll<HTMLElement>('button, [tabindex="0"]') ?? [])
  const first = controls[0], last = controls.at(-1)
  if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last?.focus() }
  else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first?.focus() }
}

onMounted(() => dialog.value?.showModal())
onBeforeUnmount(() => dialog.value?.close())
</script>

<template>
  <dialog ref="dialog" class="generation-details" aria-labelledby="generation-details-title" @cancel.prevent="emit('close')" @keydown="handleKey">
    <header>
      <strong id="generation-details-title">生成详情</strong>
      <button type="button" aria-label="关闭生成详情" autofocus @click="emit('close')"><i aria-hidden="true" class="ri-close-line" /></button>
    </header>
    <div class="details-body" tabindex="0" aria-label="异常详情">
      <article v-for="(entry, index) in entries" :key="index">
        <strong>{{ entry.label }}</strong>
        <p>{{ entry.message }}</p>
      </article>
    </div>
  </dialog>
</template>

<style scoped>
.generation-details { width: min(400px, calc(100vw - 28px)); max-height: calc(100vh - 100px); padding: 0; border: 1px solid #dce3e9; border-radius: 10px; background: white; color: #24303f; box-shadow: 0 12px 40px #18283b33; }
.generation-details[open] { display: flex; flex-direction: column; }
.generation-details::backdrop { background: #18283b55; }
header { display: flex; align-items: center; justify-content: space-between; gap: 12px; padding: 10px 14px; border-bottom: 1px solid #e2e7eb; }
header strong { font-size: 14px; }
button { width: 28px; height: 28px; border: 1px solid #dce3e9; border-radius: 5px; background: white; color: #53627a; cursor: pointer; font-size: 15px; }
button:hover { background: #eef6f5; }
.details-body { min-height: 0; overflow: auto; padding: 14px; display: grid; gap: 14px; font-size: 12px; overflow-wrap: anywhere; }
.details-body:focus-visible { outline: 2px solid #0aa19e; outline-offset: -2px; }
article p { margin: 4px 0 0; color: #53627a; white-space: pre-wrap; }
</style>
