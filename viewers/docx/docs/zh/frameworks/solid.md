# SolidJS

SolidJS 适配器挂载共享的 `<docx-editor>` 元素，并通过 `onDocumentChange` 转发模型变更。

```tsx
import { WordEditor } from 'docx-solid-viewer';

function Editor(props) {
	return (
		<WordEditor
			documentModel={props.model}
			onDocumentChange={props.setModel}
			editorRef={(handle) => props.setHandle(handle)}
		/>
	);
}
```

适配器会以一个句柄调用 `editorRef`，该句柄暴露 `load(bytes)` 和 `save()`。参见[完整的绑定契约](/zh/bindings)或[试用 Solid 演示](/demo-solid/)。
