// @vitest-environment happy-dom
import { afterEach, describe, expect, it } from 'vitest';

import { spaceBeforeCaretRange } from './inline-trailing-space';

afterEach(() => {
	document.getSelection()?.removeAllRanges();
	document.body.replaceChildren();
});

function caretIn(text: string, offset: number, end = offset): Text {
	const host = document.createElement('div');
	host.textContent = text;
	document.body.append(host);
	const node = host.firstChild as Text;
	const range = document.createRange();
	range.setStart(node, offset);
	range.setEnd(node, end);
	document.getSelection()?.removeAllRanges();
	document.getSelection()?.addRange(range);
	return node;
}

describe('spaceBeforeCaretRange', () => {
	it('covers the one space right before a collapsed caret', () => {
		const node = caretIn('fox jumps', 4);
		const range = spaceBeforeCaretRange(document);
		expect(range?.toString()).toBe(' ');
		range?.deleteContents();
		expect(node.data).toBe('foxjumps');
	});

	it('is null without a space before the caret, at the start, or for a selection', () => {
		caretIn('fox jumps', 3);
		expect(spaceBeforeCaretRange(document)).toBeNull();
		caretIn(' fox', 0);
		expect(spaceBeforeCaretRange(document)).toBeNull();
		caretIn('fox jumps', 4, 6);
		expect(spaceBeforeCaretRange(document)).toBeNull();
	});
});
