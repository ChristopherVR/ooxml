// @vitest-environment jsdom
import { createComponent, createRoot } from 'solid-js';
import { render } from 'solid-js/web';
import { afterEach, expect, it, vi } from 'vitest';
import { DocumentCanvas, createViewerBuildingBlocks } from './building-blocks';
import { stubSpeech } from '../../bindings/src/speech-stub';
import { createDocument } from '../../core/src/index';

afterEach(() => {
	document.body.replaceChildren();
	vi.unstubAllGlobals();
});

it('exposes Read Aloud state through the accessor and stops on dispose', async () => {
	const speech = stubSpeech();
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text: 'Solid words' }] }];
	const host = document.createElement('div');
	document.body.append(host);
	let blocks!: ReturnType<typeof createViewerBuildingBlocks>;
	let dispose = () => {};
	createRoot((stop) => {
		dispose = stop;
		render(() => {
			blocks = createViewerBuildingBlocks(() => ({ documentModel: model }));
			return createComponent(DocumentCanvas, blocks.canvasProps) as Node;
		}, host);
	});
	try {
		await vi.waitFor(() => expect(blocks.toolbarProps()).not.toBeNull());
		expect(blocks.toolbarProps()).toMatchObject({ canReadAloud: true, readingAloud: false });
		expect(blocks.toolbarProps()!.onReadAloud()).toBe('started');
		expect(speech.spoken).toEqual(['Solid words']);
		expect(blocks.toolbarProps()?.readingAloud).toBe(true);
	} finally {
		dispose();
	}
	expect(speechSynthesis.cancel).toHaveBeenCalled();
});
