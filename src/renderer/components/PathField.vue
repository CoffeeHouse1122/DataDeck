<script lang="ts">
import { computed, defineComponent } from 'vue'

export default defineComponent({
  name: 'PathField',
  props: {
    label: {
      type: String,
      required: true
    },
    description: {
      type: String,
      required: true
    },
    value: {
      type: String,
      required: true
    },
    defaultValue: {
      type: String,
      default: ''
    },
    icon: {
      type: String,
      required: true
    }
  },
  emits: {
    pick: () => true,
    reveal: () => true,
    reset: () => true
  },
  setup(props, { emit }) {
    const displayValue = computed(() => {
      if (!props.value) {
        return '点击选择文件或目录'
      }
      const parts = props.value.split(/[\\/]/).filter(Boolean)
      return parts.at(-1) ?? props.value
    })

    const isDefault = computed(() => Boolean(props.defaultValue) && props.value === props.defaultValue)

    function pick(): void {
      emit('pick')
    }

    function reveal(): void {
      emit('reveal')
    }

    function reset(): void {
      emit('reset')
    }

    return {
      displayValue,
      isDefault,
      pick,
      reveal,
      reset
    }
  }
})
</script>

<template>
  <div class="path-field">
    <div class="path-field__head">
      <div class="path-field__label">
        <i :class="icon" />
        <span>{{ label }}</span>
      </div>
      <div class="path-field__tools">
        <button v-if="value" type="button" class="ghost" title="打开所在位置" @click="reveal">
          <i class="ri-folder-open-line" />
        </button>
        <button v-if="value" type="button" class="ghost" :title="defaultValue ? '恢复默认路径' : '清除路径'" @click="reset">
          <i :class="defaultValue ? 'ri-restart-line' : 'ri-close-line'" />
        </button>
      </div>
    </div>
    <p class="path-field__desc">{{ description }}</p>
    <button type="button" class="path-field__picker" @click="pick">
      <span class="path-field__value" :title="value">{{ displayValue }}</span>
      <span v-if="isDefault" class="path-field__badge">默认</span>
      <span class="path-field__action">浏览</span>
    </button>
  </div>
</template>

<style scoped>
.path-field {
  min-width: 0;
  min-height: 98px;
  padding: 8px;
  border: 1px solid #d0d7de;
  border-radius: 8px;
  background: #fff;
  display: grid;
  grid-template-rows: 22px 26px 32px;
  gap: 5px;
}

.path-field__head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 8px;
  min-height: 22px;
}

.path-field__label {
  display: flex;
  align-items: center;
  gap: 7px;
  min-width: 0;
  font-size: 11px;
  font-weight: 700;
  color: #1f2328;
}

.path-field__label span {
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
}

.path-field__tools {
  flex: none;
  display: inline-flex;
  align-items: center;
  gap: 8px;
}

.path-field__desc {
  margin: 0;
  color: #57606a;
  font-size: 10px;
  line-height: 1.2;
  overflow: hidden;
  display: -webkit-box;
  -webkit-line-clamp: 2;
  -webkit-box-orient: vertical;
}

.path-field__picker {
  width: 100%;
  min-width: 0;
  height: 32px;
  min-height: 32px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 6px;
  border: 1px solid #d0d7de;
  background: #f6f8fa;
  border-radius: 6px;
  padding: 7px 8px;
  color: #1f2328;
  cursor: pointer;
  text-align: left;
}

.path-field__value {
  min-width: 0;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
  font-size: 10px;
}

.path-field__badge {
  flex: none;
  height: 18px;
  display: inline-flex;
  align-items: center;
  padding: 0 5px;
  border-radius: 999px;
  background: #ddf4ff;
  color: #0aa19e;
  font-size: 9px;
  font-weight: 700;
  line-height: 1;
}

.path-field__action {
  flex: none;
  white-space: nowrap;
  font-size: 10px;
  font-weight: 700;
  color: #0aa19e;
}

.ghost {
  flex: none;
  border: 1px solid #d0d7de;
  background: #fff;
  color: #57606a;
  border-radius: 6px;
  width: 24px;
  height: 24px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  cursor: pointer;
}

.ghost:hover,
.path-field__picker:hover {
  border-color: #8c959f;
  background: #f3f4f6;
}
</style>
