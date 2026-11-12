# Angular

`visio-angular-viewer` 将共享的 `<visio-viewer>` 以原生 Angular 集成的形式提供给你。

```bash
npm install visio-angular-viewer @angular/core
```

```ts
import { Component, Input } from '@angular/core';
import { VisioViewerComponent } from 'visio-angular-viewer';
import type { VisioDocument } from 'visio-angular-viewer';

@Component({
	selector: 'app-diagram',
	imports: [VisioViewerComponent],
	template: '<visio-viewer-host [document]="diagram" (documentError)="onError($event)" />',
	styles: ['visio-viewer-host { display: block; height: 600px }'],
})
export class Diagram {
	@Input() diagram: VisioDocument | null = null;
	onError(error: Error) {
		console.error(error);
	}
}
```

输出包括 `documentLoad`、`documentChange`、`documentError`、`pageChange`、`zoomChange`、`shapeSelect`、`selectionChange`、`ribbonAddIn`、`collaborationChange` 和 `fileCommand`。`fileCommand` 仅报告文件命令：若要取消该命令，请从 `events['file-command']` 回调中返回 `false`。完整的 `events` 回调映射同样可以作为输入项传入。

各框架中的属性（`document`、`pageIndex`、`zoom`、`showToolbar`、`canEdit`、`readOnly`、`ribbonAddIns`、`collaboration`、`storage`、`fileProviders`、`customization`、`events`）和句柄完全相同：请参阅[绑定约定](/zh/bindings)和[查看器 API](/zh/api)。如需改用查看器的各个部件组合自己的界面，请参阅[构建模块](/zh/building-blocks)；如需共享图纸，请参阅[协作](/zh/collaboration)。[试用 Angular 演示](/demo-angular/){target="_self"}。
