// @vitest-environment happy-dom
import { findTextHighlightsByElement } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import { renderToStaticMarkup } from 'react-dom/server';
import { expect, it } from 'vitest';

import { renderTextSegments } from './text-paragraph-render';

const element = {
	id: 'title',
	type: 'text',
	x: 0,
	y: 0,
	width: 400,
	height: 60,
	text: 'Alpha beta alpha',
	textSegments: [{ text: 'Alpha beta alpha', style: {} }],
} as PptxElement;

const match = (startOffset: number) => ({
	slideIndex: 0,
	elementId: 'title',
	segmentIndex: 0,
	startOffset,
	length: 5,
});

it('marks the matched run text, the current match apart (shared find-text-highlight)', () => {
	const highlights = findTextHighlightsByElement([match(0), match(11)], 1, 0).get('title');
	const host = document.createElement('div');
	host.innerHTML = renderToStaticMarkup(
		<>{renderTextSegments(element, '#000000', undefined, highlights)}</>,
	);
	const marks = [...host.querySelectorAll('mark[data-pptx-find-match]')].map((mark) => [
		mark.getAttribute('data-pptx-find-match'),
		mark.textContent,
	]);
	expect(marks).toEqual([
		['match', 'Alpha'],
		['current', 'alpha'],
	]);
	expect(host.textContent).toBe('Alpha beta alpha');
});
