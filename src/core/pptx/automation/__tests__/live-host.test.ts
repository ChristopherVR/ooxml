import { describe, expect, it } from 'vitest';

import { createLiveToolHost } from '../../../automation/live';
import { PptxHandler, type PptxData } from '../../index';
import { pptxTools } from '../catalogue';
import { createTestPptxBytes } from './helpers/create-test-pptx';

/** A PowerPoint editor over the real handler: every save restamps the revision and date. */
async function deckEditor() {
	let handler = new PptxHandler();
	let deck: PptxData = await handler.load(new Uint8Array(await createTestPptxBytes(2)).buffer);
	return {
		format: 'pptx' as const,
		deck: () => deck,
		save: () => handler.save(deck.slides),
		async load(bytes: Uint8Array) {
			handler = new PptxHandler();
			deck = await handler.load(new Uint8Array(bytes).buffer);
		},
	};
}

describe('the live tool host on a PowerPoint deck', () => {
	it('edits the deck again and again', async () => {
		const editor = await deckEditor();
		const host = createLiveToolHost(pptxTools, {
			// A date stamp has a resolution of a second, so a review this long changes it.
			review: () => new Promise((resolve) => setTimeout(() => resolve(true), 1100)),
		});
		host.register('deck', editor);
		const elementId = editor.deck().slides[0]?.elements[0]?.id;

		for (const x of [11, 22]) {
			const edit = await host.call('update_element', {
				target: 'deck',
				params: { slideIndex: 0, elementId, x },
			});
			expect(edit.changed).toBe(true);
			expect(editor.deck().slides[0]?.elements[0]?.x).toBe(x);
		}
	});

	it('refuses an edit when the user changed the deck meanwhile', async () => {
		const editor = await deckEditor();
		const host = createLiveToolHost(pptxTools, {
			review: () => {
				const element = editor.deck().slides[1]?.elements[0];
				if (element) element.x = 345;
				return true;
			},
		});
		host.register('deck', editor);
		const elementId = editor.deck().slides[0]?.elements[0]?.id;

		await expect(
			host.call('update_element', { target: 'deck', params: { slideIndex: 0, elementId, x: 9 } }),
		).rejects.toThrow(/changed while the edit was prepared/u);
		expect(editor.deck().slides[0]?.elements[0]?.x).not.toBe(9);
	});
});
