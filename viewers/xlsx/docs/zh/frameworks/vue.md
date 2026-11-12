# Vue

Vue 3 适配器是一个带有共享属性和短横线命名事件的组件。

```vue
<script setup lang="ts">
import { shallowRef } from 'vue';
import { createWorkbook, SpreadsheetEditor } from 'xlsx-vue-viewer';

const workbook = shallowRef(createWorkbook());
const editor = shallowRef();
</script>

<template>
	<SpreadsheetEditor
		ref="editor"
		:workbook="workbook"
		file-name="Budget.xlsx"
		@workbook-change="workbook = $event"
		@selection-change="(s) => console.log(s.ref)"
	/>
</template>
```

工作簿使用 `shallowRef`：它是一个较大的普通对象，深层响应式不会带来任何好处。模板引用公开了[句柄](/zh/bindings#handle)。

参见[完整的绑定契约](/zh/bindings)，或[试用 Vue 演示](/demo-vue/)。
