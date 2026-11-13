// @vitest-environment jsdom
import { afterEach, expect, it, vi } from 'vitest';
import { createViewerBuildingBlocks, mountDocumentCanvas, mountToolbar } from './building-blocks';
import { stubSpeech } from '../../bindings/src/speech-stub';
import { createDocument } from '../../core/src/index';

afterEach(() => {
	document.body.replaceChildren();
	vi.unstubAllGlobals();
});

it('exposes Read Aloud state and a toggle control, in a read-only document', () => {
	const speech = stubSpeech();
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text: 'Read me' }] }];
	const blocks = createViewerBuildingBlocks({ documentModel: model });
	const canvas = document.createElement('div');
	const toolbar = document.createElement('div');
	document.body.append(canvas, toolbar);
	const stops = [mountDocumentCanvas(canvas, blocks), mountToolbar(toolbar, blocks)];
	try {
		expect(blocks.state.mode).toBe('view');
		const button = toolbar.querySelector<HTMLButtonElement>('[data-control="readAloud"]')!;
		expect(button.disabled).toBe(false);
		expect(button.getAttribute('aria-pressed')).toBe('false');
		button.click();
		expect(speech.spoken).toEqual(['Read me']);
		expect(blocks.state.toolbarProps?.readingAloud).toBe(true);
		expect(toolbar.querySelector('[data-control="readAloud"]')?.getAttribute('aria-pressed')).toBe(
			'true',
		);
		speech.finish();
		expect(blocks.state.toolbarProps?.readingAloud).toBe(false);
		blocks.state.toolbarProps!.onReadAloud();
		blocks.destroy();
		expect(speechSynthesis.cancel).toHaveBeenCalled();
	} finally {
		for (const stop of stops.reverse()) stop();
		blocks.destroy();
	}
});

it('forwards readAloud.stopOnEdit to the editor and clears it', () => {
	const model = createDocument();
	const blocks = createViewerBuildingBlocks({
		documentModel: model,
		readAloud: { stopOnEdit: true },
	});
	const canvas = document.createElement('div');
	document.body.append(canvas);
	const stop = mountDocumentCanvas(canvas, blocks);
	try {
		const editor = canvas.querySelector('docx-editor') as unknown as { readAloud: unknown };
		expect(editor.readAloud).toEqual({ stopOnEdit: true });
		blocks.update({ documentModel: model, readAloud: null });
		expect(editor.readAloud).toBeNull();
	} finally {
		stop();
		blocks.destroy();
	}
});
