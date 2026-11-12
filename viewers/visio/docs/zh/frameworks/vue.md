# Vue 3

`visio-vue-viewer` 将共享的 `<visio-viewer>` 以原生 Vue 3 集成的形式提供给你。

```bash
npm install visio-vue-viewer vue
```

```vue
<script setup lang="ts">
import { VisioViewer } from 'visio-vue-viewer';
import type { VisioDocument } from 'visio-vue-viewer';
defineProps<{ diagram: VisioDocument }>();
</script>

<template>
	<VisioViewer :document="diagram" style="height: 600px" @document-error="console.error" />
</template>
```

组件的 ref 会暴露相同的句柄，并且同样会触发原生的 kebab-case 事件。对于大型的不可变文档，请使用 `shallowRef` 或 `markRaw`：将文档包裹在新的响应式代理中会改变其标识，并请求外部替换。

各框架中的属性（`document`、`pageIndex`、`zoom`、`showToolbar`、`canEdit`、`readOnly`、`ribbonAddIns`、`collaboration`、`storage`、`fileProviders`、`customization`、`events`）和句柄完全相同：请参阅[绑定约定](/zh/bindings)和[查看器 API](/zh/api)。如需改用查看器的各个部件组合自己的界面，请参阅[构建模块](/zh/building-blocks)；如需共享图纸，请参阅[协作](/zh/collaboration)。[试用 Vue 3 演示](/demo-vue/){target="_self"}。
