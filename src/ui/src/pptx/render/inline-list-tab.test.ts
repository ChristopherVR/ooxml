// @vitest-environment happy-dom
import type { PptxElement, TextSegment } from 'ooxml-core/pptx';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { attachInlineListController, handleInlineEditorTab } from './inline-list-controller';
import { initializeInlineListDom } from './inline-list-dom';
import { createInlineListModelObserver } from './inline-list-model';
import { createInlineListSeed } from './inline-list-seed';
import { applyInlineListTabIndent, MAX_BULLET_LIST_LEVEL } from './inline-list-tab';
import { buildInlineTextCommitPatch } from './inline-text-commit';

function seg(text: string, extra: Partial<TextSegment> = {}): TextSegment {
	return { text, style: {}, ...extra };
}

function paraBreak(): TextSegment {
	return { text: '\n', style: {}, isParagraphBreak: true };
}

function bulletedElement(textSegments: TextSegment[]): PptxElement {
	return {
		id: 'box-1',
		type: 'text',
		x: 0,
		y: 0,
		width: 400,
		height: 200,
		text: '',
		textStyle: { fontSize: 18 },
		textSegments,
	} as PptxElement;
}

function setSelection(
	startNode: Node,
	startOffset: number,
	endNode = startNode,
	endOffset = startOffset,
): void {
	const selection = window.getSelection();
	if (!selection) {
		throw new Error('Selection unavailable');
	}
	const range = document.createRange();
	range.setStart(startNode, startOffset);
	range.setEnd(endNode, endOffset);
	selection.removeAllRanges();
	selection.addRange(range);
}

function textNodeInBlock(root: HTMLElement, blockIndex: number): Text {
	const block = root.children[blockIndex];
	if (!block) {
		throw new Error(`Missing block ${blockIndex}`);
	}
	const walker = document.createTreeWalker(block, NodeFilter.SHOW_TEXT);
	const node = walker.nextNode();
	if (!node || node.nodeType !== Node.TEXT_NODE) {
		throw new Error(`Missing text node in block ${blockIndex}`);
	}
	return node as Text;
}

afterEach(() => {
	document.body.innerHTML = '';
	document.head.innerHTML = '';
	vi.restoreAllMocks();
});

describe('applyInlineListTabIndent', () => {
	it('demotes a bulleted paragraph when caret is at paragraph start', () => {
		const segments: TextSegment[] = [
			seg('First', { bulletInfo: { char: '•' }, paragraphLevel: 0 }),
			paraBreak(),
			seg('Second', { bulletInfo: { char: '•' }, paragraphLevel: 0 }),
		];
		// Start of second paragraph: "First" (5) + "\n" (1) = 6
		const result = applyInlineListTabIndent(segments, { start: 6, end: 6 }, 'demote');
		expect(result.handled).toBe(true);
		if (!result.handled) {
			return;
		}
		expect(result.changed).toBe(true);
		expect(result.textSegments[0]?.paragraphLevel).toBe(0);
		expect(result.textSegments[2]?.paragraphLevel).toBe(1);
	});

	it('does not demote when caret is in the middle or end of a bulleted paragraph', () => {
		const segments: TextSegment[] = [
			seg('First', { bulletInfo: { char: '•' }, paragraphLevel: 0 }),
		];
		expect(applyInlineListTabIndent(segments, { start: 2, end: 2 }, 'demote')).toStrictEqual({
			handled: false,
		});
		expect(applyInlineListTabIndent(segments, { start: 5, end: 5 }, 'demote')).toStrictEqual({
			handled: false,
		});
	});

	it('does not demote on a single-paragraph non-collapsed selection', () => {
		const segments: TextSegment[] = [
			seg('Hello world', { bulletInfo: { char: '•' }, paragraphLevel: 0 }),
		];
		expect(applyInlineListTabIndent(segments, { start: 0, end: 5 }, 'demote')).toStrictEqual({
			handled: false,
		});
	});

	it('does not demote a non-bulleted paragraph even at paragraph start', () => {
		const segments: TextSegment[] = [
			seg('Plain line'),
			paraBreak(),
			seg('Bullet item', { bulletInfo: { char: '•' } }),
		];
		expect(applyInlineListTabIndent(segments, { start: 0, end: 0 }, 'demote')).toStrictEqual({
			handled: false,
		});
	});

	it('promotes a bulleted paragraph on Shift+Tab regardless of caret position', () => {
		const segments: TextSegment[] = [
			seg('Nested', { bulletInfo: { char: '•' }, paragraphLevel: 2 }),
		];
		const result = applyInlineListTabIndent(segments, { start: 4, end: 4 }, 'promote');
		expect(result.handled).toBe(true);
		if (!result.handled) {
			return;
		}
		expect(result.changed).toBe(true);
		expect(result.textSegments[0]?.paragraphLevel).toBe(1);
	});

	it('clamps list levels to 0..8 and still reports handled at bounds', () => {
		const atMax: TextSegment[] = [
			seg('Deep', { bulletInfo: { char: '•' }, paragraphLevel: MAX_BULLET_LIST_LEVEL }),
		];
		const demoteAtMax = applyInlineListTabIndent(atMax, { start: 0, end: 0 }, 'demote');
		expect(demoteAtMax.handled).toBe(true);
		if (demoteAtMax.handled) {
			expect(demoteAtMax.changed).toBe(false);
			expect(demoteAtMax.textSegments[0]?.paragraphLevel).toBe(MAX_BULLET_LIST_LEVEL);
		}

		const atMin: TextSegment[] = [seg('Top', { bulletInfo: { char: '•' }, paragraphLevel: 0 })];
		const promoteAtMin = applyInlineListTabIndent(atMin, { start: 1, end: 1 }, 'promote');
		expect(promoteAtMin.handled).toBe(true);
		if (promoteAtMin.handled) {
			expect(promoteAtMin.changed).toBe(false);
			expect(promoteAtMin.textSegments[0]?.paragraphLevel).toBe(0);
		}
	});

	it('demotes and promotes only bulleted paragraphs in a multi-paragraph selection', () => {
		const segments: TextSegment[] = [
			seg('Title'),
			paraBreak(),
			seg('One', { bulletInfo: { char: '•' }, paragraphLevel: 0 }),
			paraBreak(),
			seg('Two', { bulletInfo: { char: '•' }, paragraphLevel: 1 }),
		];
		// Select from "Title" (offset 2) through "Two" (offset 12)
		const demoted = applyInlineListTabIndent(segments, { start: 2, end: 12 }, 'demote');
		expect(demoted.handled).toBe(true);
		if (!demoted.handled) {
			return;
		}
		expect(demoted.changed).toBe(true);
		expect(demoted.textSegments[0]?.paragraphLevel).toBeUndefined();
		expect(demoted.textSegments[2]?.paragraphLevel).toBe(1);
		expect(demoted.textSegments[4]?.paragraphLevel).toBe(2);

		const promoted = applyInlineListTabIndent(
			demoted.textSegments,
			{ start: 2, end: 12 },
			'promote',
		);
		expect(promoted.handled).toBe(true);
		if (!promoted.handled) {
			return;
		}
		expect(promoted.changed).toBe(true);
		expect(promoted.textSegments[0]?.paragraphLevel).toBeUndefined();
		expect(promoted.textSegments[2]?.paragraphLevel).toBe(0);
		expect(promoted.textSegments[4]?.paragraphLevel).toBe(1);
	});

	it('recomputes auto-number ordinals when a numbered paragraph changes level', () => {
		const segments: TextSegment[] = [
			seg('1.', {
				bulletInfo: { autoNumType: 'arabicPeriod', autoNumStartAt: 1, paragraphIndex: 0 },
				paragraphLevel: 0,
			}),
			seg('First'),
			paraBreak(),
			seg('2.', {
				bulletInfo: { autoNumType: 'arabicPeriod', autoNumStartAt: 1, paragraphIndex: 1 },
				paragraphLevel: 0,
			}),
			seg('Second'),
			paraBreak(),
			seg('3.', {
				bulletInfo: { autoNumType: 'arabicPeriod', autoNumStartAt: 1, paragraphIndex: 2 },
				paragraphLevel: 0,
			}),
			seg('Third'),
		];
		// Caret at start of "Second" (offset 6)
		const result = applyInlineListTabIndent(segments, { start: 6, end: 6 }, 'demote');
		expect(result.handled).toBe(true);
		if (!result.handled) {
			return;
		}
		expect(result.changed).toBe(true);
		expect(result.textSegments[0]?.bulletInfo?.paragraphIndex).toBe(0);
		expect(result.textSegments[3]?.paragraphLevel).toBe(1);
		expect(result.textSegments[3]?.bulletInfo?.paragraphIndex).toBe(0);
		expect(result.textSegments[3]?.text).toBe('1.');
		expect(result.textSegments[6]?.paragraphLevel).toBe(0);
		expect(result.textSegments[6]?.bulletInfo?.paragraphIndex).toBe(1);
		expect(result.textSegments[6]?.text).toBe('2.');
	});
});

describe('handleInlineEditorTab (DOM integration)', () => {
	it('demotes on Tab at paragraph start, promotes on Shift+Tab, and preserves commit patch', () => {
		const element = bulletedElement([
			seg('First', { bulletInfo: { char: '•' }, paragraphLevel: 0 }),
			paraBreak(),
			seg('Second', { bulletInfo: { char: '•' }, paragraphLevel: 0 }),
		]);
		const seed = createInlineListSeed(element)!;
		const root = document.createElement('div');
		root.contentEditable = 'true';
		document.body.appendChild(root);
		expect(initializeInlineListDom(root, seed)).toBe(true);

		const observer = createInlineListModelObserver(element);
		const onFormat = vi.fn();
		const controller = attachInlineListController(root, seed, { onFormat });

		const secondText = textNodeInBlock(root, 1);
		setSelection(secondText, 0);

		expect(handleInlineEditorTab(root, false)).toBe('indent');
		// Native Tab indent must not trigger external onFormat so modelObserver stays current.
		expect(onFormat).not.toHaveBeenCalled();

		const afterDemote = controller.read();
		expect(afterDemote.kind).toBe('supported');
		if (afterDemote.kind !== 'supported') {
			return;
		}
		expect(afterDemote.snapshot.text).toBe('First\nSecond');
		expect(afterDemote.snapshot.textSegments?.[2]?.paragraphLevel).toBe(1);
		expect(observer.check(element, afterDemote)).toStrictEqual({ kind: 'current' });

		// Commit patch captures the changed paragraphLevel even though plain text is unchanged.
		const patch = buildInlineTextCommitPatch(
			element,
			afterDemote.snapshot.text,
			afterDemote.snapshot,
		);
		expect(patch).toBeDefined();
		expect((patch as { textSegments?: TextSegment[] }).textSegments?.[2]?.paragraphLevel).toBe(1);

		// Shift+Tab mid-paragraph promotes back to level 0.
		const updatedText = textNodeInBlock(root, 1);
		setSelection(updatedText, 3);
		expect(handleInlineEditorTab(root, true)).toBe('indent');

		const afterPromote = controller.read();
		expect(afterPromote.kind).toBe('supported');
		if (afterPromote.kind === 'supported') {
			expect(afterPromote.snapshot.textSegments?.[2]?.paragraphLevel).toBe(0);
		}

		controller.dispose();
	});

	it('inserts a tab character when Tab is pressed mid-paragraph in a bulleted item', () => {
		const element = bulletedElement([
			seg('Hello', { bulletInfo: { char: '•' }, paragraphLevel: 0 }),
		]);
		const seed = createInlineListSeed(element)!;
		const root = document.createElement('div');
		root.contentEditable = 'true';
		document.body.appendChild(root);
		expect(initializeInlineListDom(root, seed)).toBe(true);

		const controller = attachInlineListController(root, seed);
		const textNode = textNodeInBlock(root, 0);
		setSelection(textNode, 2);

		expect(handleInlineEditorTab(root, false)).toBe('tab');
		const afterTab = controller.read();
		expect(afterTab.kind).toBe('supported');
		if (afterTab.kind === 'supported') {
			expect(afterTab.snapshot.text).toBe('He\tllo');
			expect(afterTab.snapshot.textSegments?.[0]?.paragraphLevel).toBe(0);
		}
		controller.dispose();
	});

	it('falls back to tab insertion and Shift+Tab suppression in plain non-list editors', () => {
		const root = document.createElement('div');
		root.contentEditable = 'true';
		root.textContent = 'Plain';
		document.body.appendChild(root);
		const node = root.firstChild as Text;
		setSelection(node, 0);

		expect(handleInlineEditorTab(root, true)).toBe('suppress');
		expect(root.textContent).toBe('Plain');

		expect(handleInlineEditorTab(root, false)).toBe('tab');
		expect(root.textContent).toBe('\tPlain');
	});
});
