// @vitest-environment jsdom
import '@angular/compiler';
import { Component, createComponent, inject, provideZonelessChangeDetection } from '@angular/core';
import { createApplication } from '@angular/platform-browser';
import { expect, it } from 'vitest';
import {
	DocumentCanvasComponent,
	PageNavigationComponent,
	StatusBarComponent,
	ToolbarComponent,
	ViewerBuildingBlocks,
} from './building-blocks';
import { createDocument } from '../../core/src/index';

@Component({
	standalone: true,
	imports: [DocumentCanvasComponent, ToolbarComponent, PageNavigationComponent, StatusBarComponent],
	providers: [ViewerBuildingBlocks],
	template: `
		<word-document-canvas [canvasProps]="blocks.canvasProps" />
		@if (blocks.toolbarProps(); as toolbar) {
			<word-toolbar [toolbarProps]="toolbar" />
		}
		@if (blocks.navigationProps(); as navigation) {
			<word-page-navigation [navigationProps]="navigation" />
		}
		@if (blocks.statusProps(); as status) {
			<word-status-bar [statusProps]="status" />
		}
	`,
})
class Owner {
	readonly blocks = inject(ViewerBuildingBlocks);
}

it('draws the toolbar, page navigation and status controls from signals and keeps their nodes', async () => {
	const host = document.createElement('div');
	document.body.append(host);
	const app = await createApplication({ providers: [provideZonelessChangeDetection()] });
	const component = createComponent(Owner, {
		environmentInjector: app.injector,
		hostElement: host,
	});
	app.attachView(component.hostView);
	try {
		const model = createDocument();
		model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text: 'Angular composed words' }] }];
		const { blocks } = component.instance;
		blocks.update({ documentModel: model, canEdit: true });
		component.changeDetectorRef.detectChanges();
		component.changeDetectorRef.detectChanges();
		const bold = host.querySelector<HTMLButtonElement>('button[aria-label="Bold"]')!;
		expect(bold.disabled).toBe(false);
		for (const name of ['Font', 'Font size', 'Align right', 'Bulleted list', 'Zoom'])
			expect(host.querySelector(`[aria-label="${name}"]`), name).not.toBeNull();
		expect(host.textContent).toContain('Page 1 of 1');
		expect(host.textContent).toContain('3 words');
		bold.click();
		component.changeDetectorRef.detectChanges();
		expect(host.querySelector('button[aria-label="Bold"]')).toBe(bold);
		expect(bold.getAttribute('aria-pressed')).toBe('true');
		blocks.update({ documentModel: model, canEdit: false });
		component.changeDetectorRef.detectChanges();
		expect(bold.disabled).toBe(true);
		expect(blocks.mode()).toBe('view');
	} finally {
		component.destroy();
		app.destroy();
		host.remove();
	}
});
