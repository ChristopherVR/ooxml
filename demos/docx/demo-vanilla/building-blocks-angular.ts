import { Component, InjectionToken, inject, signal } from '@angular/core';
import type { BuildingBlockOptions } from '../../../viewers/docx/packages/bindings/src/building-blocks';
import {
	DocumentCanvasComponent,
	PageNavigationComponent,
	StatusBarComponent,
	ToolbarComponent,
	ViewerBuildingBlocks,
	createViewerBuildingBlocks,
} from '../../../viewers/docx/packages/angular/src/building-blocks';

/** The composition's first options, provided by the application that mounts it. */
export const INITIAL_OPTIONS = new InjectionToken<BuildingBlockOptions>('initial options');

/**
 * The Angular native composition: a top-level decorated class, so the demo build compiles it
 * ahead of time (the published site's script policy never allows the run-time compiler).
 */
@Component({
	selector: 'word-native-composition',
	standalone: true,
	imports: [DocumentCanvasComponent, ToolbarComponent, PageNavigationComponent, StatusBarComponent],
	providers: [ViewerBuildingBlocks],
	template:
		'<word-document-canvas [class]="blocks.mode()" [canvasProps]="blocks.canvasProps" /> @if (blocks.toolbarProps(); as toolbar) { <word-toolbar [toolbarProps]="toolbar" /> } @if (blocks.navigationProps(); as navigation) { <word-page-navigation [navigationProps]="navigation" /> } @if (blocks.statusProps(); as status) { <word-status-bar [statusProps]="status" /> }',
})
export class NativeComposition {
	readonly options = signal(inject(INITIAL_OPTIONS));
	readonly blocks = createViewerBuildingBlocks(() => this.options());
}
