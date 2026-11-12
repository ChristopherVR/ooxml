# Vanilla JavaScript

`visio-vanilla-viewer` 将共享的 `<visio-viewer>` 以原生 JavaScript 集成的形式提供给你。

```bash
npm install visio-vanilla-viewer
```

```ts
import { mountViewer } from 'visio-vanilla-viewer';

const viewer = mountViewer(host, {
	document: diagram,
	zoom: 1,
	events: { 'document-error': console.error },
});

viewer.update({ pageIndex: 1 });
await viewer.load(file);
viewer.destroy(); // idempotent
```

`mountViewer` 返回包含 `update()` 的完整共享绑定，以及可重复调用的 `destroy()`。

各框架中的属性（`document`、`pageIndex`、`zoom`、`showToolbar`、`canEdit`、`readOnly`、`ribbonAddIns`、`collaboration`、`storage`、`fileProviders`、`customization`、`events`）和句柄完全相同：请参阅[绑定约定](/zh/bindings)和[查看器 API](/zh/api)。如需改用查看器的各个部件组合自己的界面，请参阅[构建模块](/zh/building-blocks)；如需共享图纸，请参阅[协作](/zh/collaboration)。[试用原生 JavaScript 演示](/demo/){target="_self"}。
