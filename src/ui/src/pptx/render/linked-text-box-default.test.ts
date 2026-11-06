/**
 * PowerPoint does not flow `a:linkedTxbx` (measured with COM Slide.Export on
 * the linked-textbox fixture: the head keeps all text, unclipped, successors
 * stay empty), so the shared renderer must not either by default.
 */
import type { PptxElementWithText } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import {
	buildSlideOverflowMap,
	getOverflowSegments,
	isLinkedTextBox,
} from './linked-text-box-overflow';
import { buildTextBlockStyle } from './text-block-style';

function box(id: string, seq: number, text?: string): PptxElementWithText {
	return {
		type: 'text',
		id,
		x: 0,
		y: 0,
		width: 60,
		height: 30,
		textStyle: { fontSize: 18 },
		linkedTxbxId: 7,
		linkedTxbxSeq: seq,
		...(text ? { textSegments: [{ text, style: {} }] } : {}),
	} as PptxElementWithText;
}

describe('linked text boxes by default', () => {
	it('does not flow text and does not clip the head', () => {
		const head = box('head', 0, 'ABCDEFGHIJ');
		const tail = box('tail', 1);
		expect(isLinkedTextBox(head)).toBe(false);
		expect(getOverflowSegments(head, [head, tail])).toBeUndefined();
		expect(getOverflowSegments(tail, [head, tail])).toBeUndefined();
		expect(buildSlideOverflowMap([head, tail]).size).toBe(0);
		expect(buildTextBlockStyle(head).overflow).not.toBe('hidden');
		// The flow is still reachable by asking for it.
		expect(isLinkedTextBox(head, { flow: true })).toBe(true);
		expect(getOverflowSegments(head, [head, tail], { flow: true })).toBeDefined();
		expect(buildTextBlockStyle(head, { linkedFlow: true }).overflow).toBe('hidden');
	});
});
