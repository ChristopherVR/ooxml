# Solid

`visio-solid-viewer` 将共享的 `<visio-viewer>` 以原生 Solid 集成的形式提供给你。

```bash
npm install visio-solid-viewer solid-js
```

```tsx
import { VisioViewer } from 'visio-solid-viewer';
import type { VisioDocument } from 'visio-solid-viewer';

export function Diagram(props: { diagram: VisioDocument }) {
	return (
		<VisioViewer
			document={props.diagram}
			style={{ height: '600px' }}
			events={{ 'document-error': console.error }}
		/>
	);
}
```

传入 `viewerRef` 即可在挂载时接收句柄（卸载时为 `undefined`）。

各框架中的属性（`document`、`pageIndex`、`zoom`、`showToolbar`、`canEdit`、`readOnly`、`ribbonAddIns`、`collaboration`、`storage`、`fileProviders`、`customization`、`events`）和句柄完全相同：请参阅[绑定约定](/zh/bindings)和[查看器 API](/zh/api)。如需改用查看器的各个部件组合自己的界面，请参阅[构建模块](/zh/building-blocks)；如需共享图纸，请参阅[协作](/zh/collaboration)。[试用 Solid 演示](/demo-solid/){target="_self"}。
