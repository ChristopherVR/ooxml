// @vitest-environment happy-dom
/**
 * `listConversionScope` reaches every list command in `useTextElementOperations`.
 *
 * The hook reads the caret from the live DOM selection, so the shared
 * `getInlineEditorSelectionResult` is stood in for with a bare caret in the
 * middle item. Everything else (the list conversion itself) is the production
 * path from `ooxml-ui/pptx`.
 */
import type { ListConversionScope } from 'ooxml-ui/pptx';
import type { PptxElement, TextSegment } from 'pptx-viewer-core';
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import { describe, expect, it, vi } from 'vitest';

/** A collapsed caret inside the second item's text run (segment 4). */
const CARET = { startSegIdx: 4, startOffset: 1, endSegIdx: 4, endOffset: 1 };

vi.mock(import('ooxml-ui/pptx'), async (importOriginal) => {
	const actual = await importOriginal();
	return {
		...actual,
		getInlineEditorSelectionResult: () => ({ kind: 'supported' as const, selection: CARET }),
	};
});

const { paragraphBulletKind, splitBulletParagraphs } = await import('ooxml-ui/pptx');
const { useTextElementOperations } = await import('./useTextElementOperations');

const seg = (text: string, extra: Partial<TextSegment> = {}): TextSegment => ({
	text,
	style: { fontSize: 18 },
	...extra,
});
const brk = (): TextSegment => ({ text: '\n', style: {}, isParagraphBreak: true });
const bullet = (text: string): TextSegment[] => [
	seg('•', { bulletInfo: { char: '•' }, paragraphLevel: 0 }),
	seg(text, { paragraphLevel: 0 }),
];

/** A three-item bulleted list: A, B, C. */
function bulletedList(): PptxElement {
	const segments = [...bullet('A'), brk(), ...bullet('B'), brk(), ...bullet('C')];
	return {
		type: 'text',
		id: 'list',
		x: 0,
		y: 0,
		width: 200,
		height: 100,
		text: segments.map((s) => s.text).join(''),
		textStyle: { fontSize: 18 },
		textSegments: segments,
	} as unknown as PptxElement;
}

type Command = (ops: ReturnType<typeof useTextElementOperations>) => void;

/** Run one list command against the bulleted list and return each paragraph's list kind. */
function kindsAfter(command: Command, listConversionScope?: ListConversionScope): string[] {
	const updateSelectedElement = vi.fn<(updates: Partial<PptxElement>) => void>();
	let ops: ReturnType<typeof useTextElementOperations> | undefined;
	function Harness(): null {
		ops = useTextElementOperations({
			selectedElement: bulletedList(),
			inlineEditingElementId: null,
			inlineEditingText: '',
			setInlineEditingElementId: () => {},
			updateSelectedElement,
			listConversionScope,
		});
		return null;
	}
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	const root = createRoot(document.createElement('div'));
	try {
		act(() => root.render(<Harness />));
		act(() => command(ops!));
	} finally {
		act(() => root.unmount());
		globalThis.IS_REACT_ACT_ENVIRONMENT = false;
	}
	const patch = updateSelectedElement.mock.calls[0]?.[0] as { textSegments: TextSegment[] };
	return splitBulletParagraphs(patch.textSegments).map((paragraph) =>
		paragraphBulletKind(paragraph.segments),
	);
}

const toggleNumbering: Command = (ops) => ops.toggleSelectedBullets('numbered');
const pickNumbering: Command = (ops) =>
	ops.updateSelectedTextStyle({ listType: 'numbered', autoNumType: 'arabicPeriod' });

describe('useTextElementOperations listConversionScope', () => {
	it('converts only the caret paragraph by default (PowerPoint)', () => {
		expect(kindsAfter(toggleNumbering)).toStrictEqual(['bullet', 'numbered', 'bullet']);
	});

	it("converts only the caret paragraph with 'paragraph'", () => {
		expect(kindsAfter(toggleNumbering, 'paragraph')).toStrictEqual([
			'bullet',
			'numbered',
			'bullet',
		]);
	});

	it("converts the whole list from a caret with 'list' (Numbering toggle)", () => {
		expect(kindsAfter(toggleNumbering, 'list')).toStrictEqual(['numbered', 'numbered', 'numbered']);
	});

	it("converts the whole list from a caret with 'list' (numbering style picker)", () => {
		expect(kindsAfter(pickNumbering)).toStrictEqual(['bullet', 'numbered', 'bullet']);
		expect(kindsAfter(pickNumbering, 'list')).toStrictEqual(['numbered', 'numbered', 'numbered']);
	});
});
