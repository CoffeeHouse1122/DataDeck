<script setup lang="ts">
import { computed } from 'vue'
const props = withDefaults(defineProps<{
  label: string
  description: string
  value: string
  defaultValue?: string
  icon: string
  compact?: boolean
  fullPath?: boolean
  disabled?: boolean
}>(), { defaultValue: '', compact: false, fullPath: false, disabled: false })
defineEmits<{ pick: []; reveal: []; reset: [] }>()
const displayValue = computed(() => props.value ? props.fullPath ? props.value : props.value.split(/[\\/]/).filter(Boolean).at(-1) : '点击选择')
</script>

<template>
  <div class="path-field" :class="{ compact }">
    <div class="path-label" :title="description"><span>{{ label }}</span></div>
    <div class="path-controls">
      <button class="path-picker" :title="value || description" :aria-label="'选择' + label" :disabled="disabled" @click="$emit('pick')">
        <i :class="icon" aria-hidden="true" /><span :class="{ placeholder: !value }">{{ displayValue }}</span>
      </button>
      <button class="path-tool" :title="'打开' + label + '所在位置'" :aria-label="'打开' + label + '所在位置'" :disabled="!value" @click="$emit('reveal')"><i aria-hidden="true" class="ri-folder-open-line" /></button>
      <button class="path-tool" :title="defaultValue ? '恢复默认路径' : '清除路径'" :aria-label="(defaultValue ? '重置' : '清除') + label" :disabled="disabled || !value" @click="$emit('reset')"><i aria-hidden="true" :class="defaultValue ? 'ri-restart-line' : 'ri-close-line'" /></button>
    </div>
  </div>
</template>

<style scoped>
.path-field { display: grid; gap: 8px; min-width: 0; }
.path-label { min-width: 0; font-weight: 650; font-size: 13px; color: #24303f; }
.path-controls { display: flex; gap: 6px; min-width: 0; }
.path-picker, .path-tool { height: 34px; border: 1px solid #dce3e9; border-radius: 5px; background: white; color: #354359; display: inline-flex; align-items: center; cursor: pointer; }
.path-picker { flex: 1; min-width: 0; gap: 8px; padding: 0 8px; text-align: left; font-size: 12px; }
.path-picker span { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.path-picker i { flex: none; font-size: 18px; color: #16835b; }
.path-picker .ri-file-ppt-2-line { color: #c45a35; }
.path-picker .ri-folder-3-line { color: #6f7e8f; }
.placeholder { color: #7c8999; }
.path-tool { width: 28px; justify-content: center; flex: none; padding: 0; font-size: 15px; }
button:hover:not(:disabled) { background: #f0f8f8; border-color: #9abcbf; }
button:focus-visible { outline: 2px solid #0aa19e; outline-offset: 2px; }
button:disabled { cursor: default; opacity: .45; }
.compact { grid-template-columns: 94px minmax(0, 1fr); gap: 8px; align-items: center; }
.compact .path-label { font-weight: 400; font-size: 12px; }
.compact .path-picker { gap: 5px; padding-inline: 6px; }
@media (max-width: 390px) { .compact { grid-template-columns: 80px minmax(0, 1fr); gap: 5px; } .path-controls { gap: 4px; } }
</style>
