<script lang="ts">
import { computed, defineComponent, onBeforeUnmount, onMounted, ref, type PropType } from 'vue'

type Option = {
  label: string
  value: string
}

export default defineComponent({
  name: 'CustomSelect',
  props: {
    modelValue: {
      type: String,
      required: true
    },
    options: {
      type: Array as PropType<Option[]>,
      required: true
    }
  },
  emits: {
    'update:modelValue': (value: string) => typeof value === 'string'
  },
  setup(props, { emit }) {
    const root = ref<HTMLElement | null>(null)
    const open = ref(false)

    const activeLabel = computed(() => props.options.find((option) => option.value === props.modelValue)?.label ?? '请选择')

    function toggle(): void {
      open.value = !open.value
    }

    function choose(value: string): void {
      emit('update:modelValue', value)
      open.value = false
    }

    function handleClick(event: MouseEvent): void {
      if (!root.value?.contains(event.target as Node)) {
        open.value = false
      }
    }

    onMounted(() => {
      window.addEventListener('click', handleClick)
    })

    onBeforeUnmount(() => {
      window.removeEventListener('click', handleClick)
    })

    return {
      activeLabel,
      choose,
      open,
      root,
      toggle
    }
  }
})
</script>

<template>
  <div ref="root" class="select-root">
    <button type="button" class="select-trigger" @click.stop="toggle">
      <span>{{ activeLabel }}</span>
      <i class="ri-arrow-down-s-line" />
    </button>
    <div v-if="open" class="select-menu">
      <button
        v-for="option in options"
        :key="option.value"
        type="button"
        class="select-option"
        :class="{ active: option.value === modelValue }"
        @click.stop="choose(option.value)"
      >
        <span>{{ option.label }}</span>
        <i v-if="option.value === modelValue" class="ri-check-line" />
      </button>
    </div>
  </div>
</template>

<style scoped>
.select-root {
  position: relative;
}

.select-trigger,
.select-option {
  width: 100%;
  border: 1px solid #d0d7de;
  background: #fff;
  color: #1f2328;
  padding: 10px 12px;
  border-radius: 6px;
  font-size: 13px;
  line-height: 1.2;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 10px;
}

.select-trigger {
  cursor: pointer;
}

.select-menu {
  position: absolute;
  inset: calc(100% + 6px) 0 auto 0;
  border: 1px solid #d0d7de;
  background: #fff;
  border-radius: 6px;
  box-shadow: 0 12px 28px rgba(31, 35, 40, 0.12);
  overflow: hidden;
  z-index: 20;
}

.select-option {
  border: 0;
  border-radius: 0;
  cursor: pointer;
}

.select-option:hover,
.select-option.active {
  background: #f6f8fa;
}
</style>
