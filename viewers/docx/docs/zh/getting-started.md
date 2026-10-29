# 快速开始

docx-viewer 由一个 `<docx-editor>` Web Component 和六个轻量的框架适配器组成。本页将安装适配器、挂载编辑器，并加载和保存文档。它依据[稳定性策略](https://github.com/ChristopherVR/ooxml/blob/main/docs/stability.md)属于稳定版，并提供[保留保证](/preservation)，但并不等同于 Microsoft Word（参见[支持路线图](/parity-roadmap)）。

## 1. 安装

安装适用于你所用框架的自包含包。它已打包编辑器界面和适配器，并自带 `docx-core`（文档模型）。

::: code-group

```bash [React]
npm install docx-react-viewer react
```

```bash [Vue]
npm install docx-vue-viewer vue
```

```bash [Angular]
npm install docx-angular-viewer
```

```bash [Svelte]
npm install docx-svelte-viewer svelte
```

```bash [Solid]
npm install docx-solid-viewer solid-js
```

```bash [Vanilla JS]
npm install docx-vanilla-viewer
```

:::

仅在无头场景（不使用编辑器，只做解析和序列化）下才需要单独安装 `docx-core`。

## 2. 挂载编辑器

所有适配器挂载的都是同一个 `<docx-editor>` 元素。传入一个 `DocumentModel`，监听变更，并为容器设置高度。

```tsx
import { useState } from 'react';
import { createDocument, WordEditor } from 'docx-react-viewer';

export function Editor() {
	const [model, setModel] = useState(() => createDocument());
	return (
		<div style={{ height: '100vh' }}>
			<WordEditor documentModel={model} onDocumentChange={setModel} />
		</div>
	);
}
```

其他框架使用相同的选项：参见[框架指南](/frameworks/react)和[绑定约定](/bindings)。

## 3. 加载和保存文件

组件引用（或原生 JS 的句柄）提供 `load(bytes)` 和 `save()`。加载时会识别 DOCX 和旧版 `.doc`；不支持受密码保护的文件。

```ts
const bytes = new Uint8Array(await file.arrayBuffer());
await editor.load(bytes);

// After edits:
const blob = await editor.save(); // Blob of the serialized .docx
```

`dirty` 属性和 `dirty-change` 事件用于报告未保存的编辑。所有属性、方法和事件请参见[元素 API](/zh/api)。

## 4. 亲自试用

- [在线演示](/demos)在每个框架适配器中运行同一个编辑器。
- 要在仓库克隆中运行演示（需要 Bun）：

```bash
bun install
bun run demo
```

## 后续步骤

- [元素 API 与事件](/zh/api)
- [主题与明暗模式同步](/theming)
- [编辑文本](/zh/editing)和[协作](/zh/collaboration)
- [架构](/zh/architecture)和[模型单位](/model-units)
