// @vitest-environment jsdom
import type { FindResult } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { FIND_MATCH_ATTRIBUTE, markFindHighlights, syncFindHighlights } from './find-highlight-dom';
import { mediaElementFromFile } from './media-element-from-file';

const match = (elementId: string, slideIndex = 0) =>
	({ slideIndex, elementId, segmentIndex: 0, startOffset: 0, length: 1 }) as FindResult;

function stage() {
	const root = document.createElement('div');
	root.innerHTML = '<div data-element-id="a"></div><div data-element-id="b"></div>';
	return root;
}

describe('find highlights', () => {
	it('marks the matching elements of the slide, the current one as current', () => {
		const root = stage();
		syncFindHighlights(root, [match('a'), match('b'), match('c', 1)], 1, 0);
		const mark = (id: string) =>
			root.querySelector(`[data-element-id="${id}"]`)?.getAttribute(FIND_MATCH_ATTRIBUTE);
		expect(mark('a')).toBe('match');
		expect(mark('b')).toBe('current');
		syncFindHighlights(root, [match('a')], 0, 0);
		expect(mark('a')).toBe('current');
		expect(mark('b')).toBeNull();
		markFindHighlights(root, []);
		expect(root.querySelector(`[${FIND_MATCH_ATTRIBUTE}]`)).toBeNull();
	});
});

describe('mediaElementFromFile', () => {
	it('builds a centred audio strip and refuses other files', async () => {
		const audio = new File([new Uint8Array([1, 2])], 'a.mp3', { type: 'audio/mpeg' });
		const element = await mediaElementFromFile(audio, { width: 960, height: 540 });
		expect(element).toMatchObject({ type: 'media', mediaType: 'audio', width: 420, height: 64 });
		expect(element?.x).toBe(270);
		const text = new File(['x'], 'a.txt', { type: 'text/plain' });
		expect(await mediaElementFromFile(text, { width: 960, height: 540 })).toBeNull();
	});
});
