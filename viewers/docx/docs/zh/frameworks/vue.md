# Vue

Vue 适配器将编辑器暴露为一个组件，具有模型属性和 Vue 事件。

```vue
<script setup>
import { WordEditor } from 'docx-vue-viewer';
const model = defineModel();
</script>

<template>
	<WordEditor :document-model="model" @document-change="model = $event" />
</template>
```

参见[完整的绑定契约](/zh/bindings)或[试用 Vue 演示](/demo-vue/)。
