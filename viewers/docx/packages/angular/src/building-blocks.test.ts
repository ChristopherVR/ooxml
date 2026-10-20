// @vitest-environment jsdom
import '@angular/compiler';
import { Component, createComponent, inject, provideZonelessChangeDetection } from '@angular/core';
import { createApplication } from '@angular/platform-browser';
import { expect, it, vi } from 'vitest';
import { DocumentCanvasComponent, ViewerBuildingBlocks } from './building-blocks';
import { createDocument } from '../../core/src/index';
@Component({
	standalone: true,
	imports: [DocumentCanvasComponent],
	providers: [ViewerBuildingBlocks],
	template: '<word-document-canvas [canvasProps]="canvasProps" />',
})
class DirectOwner {
	readonly blocks = inject(ViewerBuildingBlocks);
	canvasProps = { ...this.blocks.canvasProps };
}
it('owns default DI events and preserves the engine when canvas props are recreated', async () => {
	const host = document.createElement('div');
	document.body.append(host);
	const app = await createApplication({ providers: [provideZonelessChangeDetection()] });
	const component = createComponent(DirectOwner, {
		environmentInjector: app.injector,
		hostElement: host,
	});
	app.attachView(component.hostView);
	try {
		component.changeDetectorRef.detectChanges();
		const blocks = component.instance.blocks;
		const engine = blocks.handle.element;
		engine.documentModel = createDocument();
		const changed = vi.fn();
		const dirty = vi.fn();
		const complete = vi.fn();
		blocks.documentChange.subscribe({ next: changed, complete });
		blocks.dirtyChange.subscribe(dirty);
		const model = createDocument();
		engine.dispatchEvent(new CustomEvent('document-change', { detail: model }));
		engine.dispatchEvent(new CustomEvent('dirty-change', { detail: true }));
		expect(changed).toHaveBeenCalledWith(model);
		expect(dirty).toHaveBeenCalledWith(true);
		blocks.toolbarProps()!.onBold();
		expect(blocks.toolbarProps()?.bold).toBe(true);
		component.instance.canvasProps = { ...blocks.canvasProps };
		component.changeDetectorRef.detectChanges();
		expect(blocks.handle.element).toBe(engine);
		expect(blocks.toolbarProps()?.bold).toBe(true);
		component.destroy();
		expect(complete).toHaveBeenCalledOnce();
		expect(host.querySelector('docx-editor')).toBeNull();
	} finally {
		if (!component.hostView.destroyed) component.destroy();
		app.destroy();
		host.remove();
	}
});
