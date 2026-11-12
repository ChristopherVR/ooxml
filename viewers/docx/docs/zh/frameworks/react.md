# React

React 适配器挂载共享的 `<docx-editor>` 元素，并通过 `onDocumentChange` 转发模型更新。

```tsx
import { WordEditor } from 'docx-react-viewer';

export function Editor({ model, setModel }) {
	return <WordEditor documentModel={model} onDocumentChange={setModel} />;
}
```

使用组件的 ref 调用 `load(bytes)` 和 `save()`。参见[完整的绑定契约](/zh/bindings)或[试用 React 演示](/demo/)。
