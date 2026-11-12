# 原生 JavaScript

导入原生挂载适配器并提供一个宿主元素。返回的句柄包括 load、save、update、destroy 以及自定义元素。

```ts
import { mountEditor } from 'docx-vanilla-viewer';

const editor = mountEditor(document.querySelector('#editor'), {
	documentModel: model,
	onDocumentChange(next) {
		model = next;
	},
});
```

也可以直接使用 `docx-vanilla-viewer` 中的 `registerDocxEditor` 注册自定义元素。参见[完整的绑定契约](/zh/bindings)或[试用原生 JavaScript 演示](/demo-vanilla/)。
