import { Component, inject } from '@angular/core';
import {
	SpreadsheetBuildingBlocks,
	SpreadsheetCanvasComponent,
	ToolbarComponent,
	FormulaBarComponent,
	SheetTabsComponent,
} from '../index';
@Component({
	standalone: true,
	providers: [SpreadsheetBuildingBlocks],
	imports: [SpreadsheetCanvasComponent, ToolbarComponent, FormulaBarComponent, SheetTabsComponent],
	template: `@if (blocks.toolbarProps(); as p) {
			<spreadsheet-toolbar [toolbarProps]="p" /><button (click)="p.onItalic()">Host italic</button>
			<spreadsheet-formula-bar [toolbarProps]="p" /><spreadsheet-sheet-tabs [toolbarProps]="p" />
		}
		<spreadsheet-canvas [mount]="blocks.canvasProps.mount" />`,
})
export class NativeBlocksHost {
	readonly blocks = inject(SpreadsheetBuildingBlocks);
}
