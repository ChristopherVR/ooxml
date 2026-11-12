// @vitest-environment jsdom
import { afterEach, expect, it, vi } from 'vitest';
import { act, createElement } from 'react';
import { createRoot } from 'react-dom/client';
import { DocumentCanvas, Toolbar, useViewerBuildingBlocks } from './building-blocks';
import { stubSpeech } from '../../bindings/src/speech-stub';
import { createDocument } from '../../core/src/index';

afterEach(() => {
	document.body.replaceChildren();
	vi.unstubAllGlobals();
});

it('exposes Read Aloud state, a control, and stops on unmount', async () => {
	vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true);
	const speech = stubSpeech();
	const model = createDocument();
	model.blocks = [{ type: 'paragraph', id: 'p', runs: [{ text: 'React words' }] }];
	const host = document.createElement('div');
	document.body.append(host);
	const root = createRoot(host);
	let blocks!: ReturnType<typeof useViewerBuildingBlocks>;
	function Host() {
		blocks = useViewerBuildingBlocks({ documentModel: model });
		return createElement(
			'div',
			null,
			createElement(DocumentCanvas, blocks.canvasProps),
			blocks.toolbarProps && createElement(Toolbar, blocks.toolbarProps),
		);
	}
	await act(async () => root.render(createElement(Host)));
	expect(blocks.mode).toBe('view');
	expect(blocks.toolbarProps).toMatchObject({ canReadAloud: true, readingAloud: false });
	await act(async () => {
		host.querySelector<HTMLButtonElement>('button[aria-label="Read aloud"]')!.click();
	});
	expect(speech.spoken).toEqual(['React words']);
	expect(blocks.toolbarProps?.readingAloud).toBe(true);
	await act(async () => {
		speech.finish();
	});
	expect(blocks.toolbarProps?.readingAloud).toBe(false);
	await act(async () => {
		blocks.toolbarProps!.onReadAloud();
	});
	await act(async () => root.unmount());
	expect(speechSynthesis.cancel).toHaveBeenCalled();
});
