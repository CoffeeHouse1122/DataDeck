<script lang="ts">
import { defineComponent } from 'vue'

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
    icon: {
      type: String,
      required: true
    }
  },
  emits: {
    pick: () => true,
    reveal: () => true
  },
  setup(_props, { emit }) {
    function pick(): void {
      emit('pick')
    }

    function reveal(): void {
      emit('reveal')
    }

    return {
      pick,
      reveal
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
      <button v-if="value" type="button" class="ghost" @click="reveal">
        <i class="ri-folder-open-line" />
      </button>
    </div>
    <p class="path-field__desc">{{ description }}</p>
    <button type="button" class="path-field__picker" @click="pick">
      <span class="path-field__value">{{ value || '点击选择文件或目录' }}</span>
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
  gap: 7px;
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

.path-field__action {
  flex: none;
  white-space: nowrap;
  font-size: 10px;
  font-weight: 700;
  color: #0969da;
}

.ghost {
  flex: none;
  border: 1px solid #d0d7de;
  background: #fff;
  color: #57606a;
  border-radius: 6px;
  width: 26px;
  height: 24px;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  cursor: pointer;
}
</style>
