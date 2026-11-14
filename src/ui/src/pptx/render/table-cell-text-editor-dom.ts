/**
 * table-cell-text-editor-dom.ts: the DOM half of the table cell text editor.
 * Renders a cell's rich segments into a contenteditable root (one block per
 * paragraph, one span per run, `<br data-cell-line-break>` for a soft break)
 * and reads text and caret offsets back in the same coordinates the segments
 * use: a paragraph or line break is one `\n`.
 *
 * Browser edits may add or drop wrappers; the reader only trusts text nodes,
 * marked line breaks and block boundaries, so a stray placeholder `<br>` or an
 * extra inline wrapper never changes the text.
 *
 * @module render/table-cell-text-editor-dom
 */
import type { TextSegment } from 'ooxml-core/pptx';
import { cellRunsFromSegments } from 'ooxml-core/pptx/editor/render/table-cell-rich-text';

import type { TableCellCss } from './table-style';
import { cellRunStyle } from './table-style';

const LINE_BREAK = 'data-cell-line-break';

function isBlock(node: Node): node is HTMLElement {
	return node.nodeType === 1 && ['DIV', 'P'].includes((node as HTMLElement).tagName);
}

/** A text position: inside a text node, or at a child index of an element. */
interface Slot {
	node: Node;
	offset: number;
}

interface Walk {
	text: string;
	/** Text nodes with the text offset where each starts. */
	texts: Array<{ node: Text; start: number }>;
	/** Paragraph blocks with the text offset where their content starts. */
	blocks: Array<{ node: HTMLElement; start: number }>;
	/** Offset of a DOM position, when it was asked for. */
	found?: number;
}

function walk(root: HTMLElement, target?: Slot): Walk {
	const result: Walk = { text: '', texts: [], blocks: [] };
	let blockCount = 0;
	const visit = (node: Node): void => {
		const children = Array.from(node.childNodes);
		children.forEach((child, index) => {
			if (target && result.found === undefined && target.node === node && target.offset === index) {
				result.found = result.text.length;
			}
			if (child.nodeType === 3) {
				const text = child as Text;
				if (target && result.found === undefined && target.node === text) {
					result.found = result.text.length + Math.min(target.offset, text.length);
				}
				result.texts.push({ node: text, start: result.text.length });
				result.text += text.data;
				return;
			}
			if (child.nodeType !== 1) {
				return;
			}
			const element = child as HTMLElement;
			if (element.tagName === 'BR') {
				if (element.hasAttribute(LINE_BREAK)) {
					result.text += '\n';
				}
				return;
			}
			if (isBlock(element)) {
				if (blockCount > 0 || result.text.length > 0) {
					result.text += '\n';
				}
				blockCount++;
				result.blocks.push({ node: element, start: result.text.length });
			}
			visit(element);
		});
		if (
			target &&
			result.found === undefined &&
			target.node === node &&
			target.offset >= children.length
		) {
			result.found = result.text.length;
		}
	};
	visit(root);
	return result;
}

/** The editor's text, breaks as `\n`. */
export function readCellEditorText(root: HTMLElement): string {
	return walk(root).text;
}

/** The text offset of a DOM position inside the editor. */
export function cellEditorOffsetOf(root: HTMLElement, node: Node, offset: number): number {
	const result = walk(root, { node, offset });
	return result.found ?? result.text.length;
}

/** The current selection as text offsets, or `undefined` when it is outside the editor. */
export function readCellEditorSelection(
	root: HTMLElement,
): { start: number; end: number } | undefined {
	const selection = root.ownerDocument.getSelection();
	if (!selection || selection.rangeCount === 0) {
		return undefined;
	}
	const range = selection.getRangeAt(0);
	if (!root.contains(range.startContainer) || !root.contains(range.endContainer)) {
		return undefined;
	}
	const start = cellEditorOffsetOf(root, range.startContainer, range.startOffset);
	const end = cellEditorOffsetOf(root, range.endContainer, range.endOffset);
	return start <= end ? { start, end } : { start: end, end: start };
}

function slotAt(root: HTMLElement, offset: number, walked: Walk): Slot {
	// Prefer a text node holding the offset; at a run boundary, the earlier run.
	for (const { node, start } of walked.texts) {
		if (offset >= start && offset <= start + node.length) {
			return { node, offset: offset - start };
		}
	}
	for (let index = walked.blocks.length - 1; index >= 0; index--) {
		const block = walked.blocks[index]!;
		if (offset >= block.start) {
			return { node: block.node, offset: 0 };
		}
	}
	return { node: root, offset: 0 };
}

/** Select text offsets `[start, end]` in the editor. */
export function setCellEditorSelection(root: HTMLElement, start: number, end = start): void {
	const selection = root.ownerDocument.getSelection();
	if (!selection) {
		return;
	}
	const walked = walk(root);
	const clamp = (value: number) => Math.max(0, Math.min(walked.text.length, value));
	const a = slotAt(root, clamp(start), walked);
	const b = slotAt(root, clamp(end), walked);
	const range = root.ownerDocument.createRange();
	range.setStart(a.node, a.offset);
	range.setEnd(b.node, b.offset);
	selection.removeAllRanges();
	selection.addRange(range);
}

/** Apply a CSS map whose values may be numbers (unitless) or strings. */
function applyCss(element: HTMLElement, css: TableCellCss): void {
	for (const [key, value] of Object.entries(css)) {
		if (value === undefined || value === null) {
			continue;
		}
		const property = key.replace(/[A-Z]/g, (ch) => `-${ch.toLowerCase()}`);
		element.style.setProperty(property, String(value));
	}
}

/** The run's own CSS, with explicit "off" values so a run can undo the cell's bold or italic. */
function runCss(segment: TextSegment, cellCss: TableCellCss | undefined): TableCellCss {
	const [run] = cellRunsFromSegments([{ ...segment, text: segment.text || ' ' }]);
	const css = run ? cellRunStyle(run, cellCss) : {};
	const style = segment.style ?? {};
	if (style.bold === false) {
		css.fontWeight = 'normal';
	}
	if (style.italic === false) {
		css.fontStyle = 'normal';
	}
	if (style.underline === false && style.strikethrough !== true) {
		css.textDecoration = 'none';
	}
	return css;
}

/** Render `segments` into `root`, replacing its content. */
export function renderCellEditorSegments(
	root: HTMLElement,
	segments: readonly TextSegment[],
	cellCss?: TableCellCss,
): void {
	const doc = root.ownerDocument;
	root.replaceChildren();
	let paragraph: HTMLElement | undefined;
	let hasContent = false;
	const close = (): void => {
		if (paragraph && !hasContent) {
			paragraph.appendChild(doc.createElement('br'));
		}
	};
	const open = (layout: TextSegment['paragraphProperties']): void => {
		paragraph = doc.createElement('div');
		paragraph.setAttribute('data-cell-paragraph', '');
		if (layout?.align) {
			const align = layout.align === 'justify' ? 'justify' : layout.align;
			paragraph.style.textAlign = align === 'center' ? 'center' : align;
		}
		root.appendChild(paragraph);
		hasContent = false;
	};
	segments.forEach((segment, index) => {
		if (segment.isParagraphBreak) {
			close();
			paragraph = undefined;
			return;
		}
		if (!paragraph) {
			open(segment.paragraphProperties);
		}
		if (segment.isLineBreak) {
			const br = doc.createElement('br');
			br.setAttribute(LINE_BREAK, '');
			paragraph!.appendChild(br);
			hasContent = true;
			return;
		}
		if (!segment.text) {
			return;
		}
		const span = doc.createElement('span');
		span.setAttribute('data-seg-idx', String(index));
		applyCss(span, runCss(segment, cellCss));
		span.textContent = segment.text;
		paragraph!.appendChild(span);
		hasContent = true;
	});
	if (!paragraph) {
		open(undefined);
	}
	close();
}

/**
 * Where the caret sits for the arrow-key boundaries: at the start or end of
 * the text, and on its first or last visual line. Lines come from layout
 * rectangles; without layout (no rectangles) the first and last paragraphs
 * stand in for them.
 */
export function readCellEditorCaret(root: HTMLElement): {
	collapsed: boolean;
	atStart: boolean;
	atEnd: boolean;
	onFirstLine: boolean;
	onLastLine: boolean;
} {
	const text = readCellEditorText(root);
	const selection = readCellEditorSelection(root);
	const collapsed = !selection || selection.start === selection.end;
	const offset = selection?.end ?? text.length;
	const firstBreak = text.indexOf('\n');
	const lastBreak = text.lastIndexOf('\n');
	let onFirstLine = firstBreak < 0 || offset <= firstBreak;
	let onLastLine = lastBreak < 0 || offset > lastBreak;
	const caretRect = rectAt(root, offset);
	const startRect = rectAt(root, 0);
	const endRect = rectAt(root, text.length);
	if (caretRect && startRect && endRect) {
		const tolerance = Math.max(2, caretRect.height / 2);
		onFirstLine = Math.abs(caretRect.top - startRect.top) < tolerance;
		onLastLine = Math.abs(caretRect.bottom - endRect.bottom) < tolerance;
	}
	return {
		collapsed,
		atStart: offset === 0,
		atEnd: offset === text.length,
		onFirstLine,
		onLastLine,
	};
}

function rectAt(root: HTMLElement, offset: number): DOMRect | undefined {
	const walked = walk(root);
	const slot = slotAt(root, offset, walked);
	const range = root.ownerDocument.createRange();
	try {
		range.setStart(slot.node, slot.offset);
		range.collapse(true);
	} catch {
		return undefined;
	}
	const rect = typeof range.getClientRects === 'function' ? range.getClientRects()[0] : undefined;
	if (rect && rect.height > 0) {
		return rect;
	}
	const element = slot.node.nodeType === 1 ? (slot.node as HTMLElement) : slot.node.parentElement;
	const fallback = element?.getBoundingClientRect?.();
	return fallback && fallback.height > 0 && slot.node.nodeType === 1 ? fallback : undefined;
}
