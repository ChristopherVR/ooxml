// @vitest-environment jsdom
import '@angular/compiler';
import { Component, createComponent, inject, provideZonelessChangeDetection } from '@angular/core';
import { createApplication } from '@angular/platform-browser';
import { afterEach, expect, it, vi } from 'vitest';
import { DocumentCanvasComponent, ViewerBuildingBlocks } from './building-blocks';
import { stubSpeech } from '../../bindings/src/speech-stub';
import { createDocument } from '../../core/src/index';

@Component({
	standalone: true,
	imports: [DocumentCanvasComponent],
	providers: [ViewerBuildingBlocks],
	template: '<word-document-canvas [canvasProps]="canvasProps" />',
})
class Owner {
	readonly blocks = inject(ViewerBuildingBlocks);
	canvasProps = { ...this.blocks.canvasProps };
}

afterEach(() => vi.unstubAllGlobals());

it('exposes Read Aloud state through the signal and stops on destroy', async () => {
	const speech = stubSpeech();
	const host = document.createElement('div');
	document.body.append(host);
	const app = await createApplication({ providers: [provideZonelessChangeDetection()] });
	const component = createComponent(Owner, {
		environmentInjector: app.injector,
		hostElement: host,
	});
	app.attachView(component.hostView);
	try {
		component.changeDetectorRef.detectChanges();
		const { blocks } = component.instance;
		const model = createDocument();
		model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text: 'Angular words' }] }];
		blocks.update({ documentModel: model });
		expect(blocks.toolbarProps()).toMatchObject({ canReadAloud: true, readingAloud: false });
		expect(blocks.toolbarProps()!.onReadAloud()).toBe('started');
		expect(speech.spoken).toEqual(['Angular words']);
		expect(blocks.toolbarProps()?.readingAloud).toBe(true);
		component.destroy();
		expect(speechSynthesis.cancel).toHaveBeenCalled();
	} finally {
		if (!component.hostView.destroyed) component.destroy();
		app.destroy();
		host.remove();
	}
});

it('forwards readAloud.stopOnEdit to the editor and clears it', async () => {
	const host = document.createElement('div');
	document.body.append(host);
	const app = await createApplication({ providers: [provideZonelessChangeDetection()] });
	const component = createComponent(Owner, {
		environmentInjector: app.injector,
		hostElement: host,
	});
	app.attachView(component.hostView);
	try {
		component.changeDetectorRef.detectChanges();
		const { blocks } = component.instance;
		const model = createDocument();
		blocks.update({ documentModel: model, readAloud: { stopOnEdit: true } });
		const editor = host.querySelector('docx-editor') as unknown as { readAloud: unknown };
		expect(editor.readAloud).toEqual({ stopOnEdit: true });
		blocks.update({ documentModel: model, readAloud: null });
		expect(editor.readAloud).toBeNull();
	} finally {
		if (!component.hostView.destroyed) component.destroy();
		app.destroy();
		host.remove();
	}
});
