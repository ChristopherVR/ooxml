# 原生 JavaScript

无需框架：`mountEditor` 会在容器中创建元素，并返回[句柄](/zh/bindings#handle)，以及 `update(options)` 和 `destroy()`。

```ts
import { mountEditor } from 'xlsx-vanilla-viewer';

const editor = mountEditor(document.querySelector('#editor')!, {
	fileName: 'Budget.xlsx',
	onWorkbookChange: (workbook) => console.log(workbook.sheets.length),
	onWorkbookError: (error) => console.error(error),
});
await editor.load(bytes, 'Budget.xlsx');
editor.update({ readOnly: true });
const blob = await editor.save();
editor.destroy();
```

也可以直接使用该元素：

```ts
import { defineXlsxEditor } from 'xlsx-vanilla-viewer';

defineXlsxEditor();
const element = document.createElement('xlsx-editor');
element.setAttribute('locale', 'de');
document.body.append(element);
element.newWorkbook();
```

有关每个属性、方法和事件，请参见[元素 API](/zh/api)。

参见[完整的绑定契约](/zh/bindings)，或[试用原生 JavaScript 演示](/demo-vanilla/)。
