<script setup lang="ts">
import type { AppUpdateState } from '../../shared/contracts'

defineProps<{ state: AppUpdateState; pending: boolean; running: boolean }>()
defineEmits<{ action: [value: 'check' | 'download' | 'install'] }>()
</script>

<template>
  <section class="app-update" aria-label="应用更新">
    <header>
      <span><i class="ri-download-cloud-2-line" aria-hidden="true" /> 应用更新</span>
      <small>v{{ state.currentVersion || '—' }}</small>
    </header>
    <p role="status" aria-live="polite">{{ state.message }}</p>
    <progress v-if="state.phase === 'downloading'" :value="state.percent ?? 0" max="100" aria-label="更新下载进度" />
    <p v-if="state.phase === 'downloaded'" class="hint">
      {{ running || state.generationRunning ? '报表生成完成后才能安装。' : '点击后会保存当前设置、关闭应用并启动安装程序。' }}
    </p>
    <button v-if="state.phase === 'available'" type="button" :disabled="pending" @click="$emit('action', 'download')">下载更新</button>
    <button v-else-if="state.phase === 'downloaded'" type="button" :disabled="pending || running || state.generationRunning" @click="$emit('action', 'install')">保存并重启安装</button>
    <button v-else type="button" :disabled="pending || !['idle', 'error'].includes(state.phase)" @click="$emit('action', 'check')">
      {{ state.phase === 'checking' ? '正在检查…' : state.phase === 'error' ? '重新检查更新' : '检查更新' }}
    </button>
  </section>
</template>

<style scoped>
.app-update { border: 1px solid #d0d7de; border-radius: 9px; padding: 14px; background: #f6f8fa; color: #24292f; }
header { display: flex; justify-content: space-between; align-items: center; gap: 12px; font-size: 13px; font-weight: 600; }
header i { color: #0969da; margin-right: 4px; }
small { color: #57606a; font-variant-numeric: tabular-nums; font-weight: 400; }
p { margin: 10px 0; font-size: 12px; line-height: 1.7; overflow-wrap: anywhere; }
.hint { color: #57606a; }
progress { width: 100%; height: 6px; accent-color: #2f81f7; display: block; margin: 12px 0; }
button { border: 1px solid var(--button-primary, #151b25); border-radius: 6px; background: var(--button-primary, #151b25); color: #fff; font: inherit; font-size: 12px; padding: 7px 12px; cursor: pointer; }
button:hover:not(:disabled) { background: var(--button-primary-hover, #263244); border-color: var(--button-primary-hover, #263244); }
button:disabled { opacity: .45; cursor: default; }
button:focus-visible { outline: 2px solid #79c0ff; outline-offset: 3px; }
</style>
