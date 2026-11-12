# React

`visio-react-viewer` 将共享的 `<visio-viewer>` 以原生 React 集成的形式提供给你。

```bash
npm install visio-react-viewer react
```

```tsx
import { useRef } from 'react';
import { VisioViewer, type ViewerHandle, type VisioDocument } from 'visio-react-viewer';

export function Diagram({ diagram }: { diagram: VisioDocument }) {
	const ref = useRef<ViewerHandle>(null);
	return (
		<VisioViewer
			ref={ref}
			document={diagram}
			style={{ height: 600 }}
			events={{ 'document-error': console.error }}
		/>
	);
}
```

转发的 `ref` 会暴露 `ViewerHandle`。React 的 `className` 和 `style` 用于设置宿主元素的样式，且宿主必须设置高度。

各框架中的属性（`document`、`pageIndex`、`zoom`、`showToolbar`、`canEdit`、`readOnly`、`ribbonAddIns`、`collaboration`、`storage`、`fileProviders`、`customization`、`events`）和句柄完全相同：请参阅[绑定约定](/zh/bindings)和[查看器 API](/zh/api)。如需改用查看器的各个部件组合自己的界面，请参阅[构建模块](/zh/building-blocks)；如需共享图纸，请参阅[协作](/zh/collaboration)。[试用 React 演示](/demo-react/){target="_self"}。
