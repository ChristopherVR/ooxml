/**
 * table-cell-text-editor.ts: the rich text editor every binding mounts into a
 * table cell being edited.
 *
 * It edits the cell's text as the shape text model does (`TextSegment`s,
 * shared with shape text through `tableCellSegments`): runs keep their
 * formatting while words in them are typed over, Enter starts a paragraph,
 * Shift+Enter a line break, and Ctrl/Cmd+B, I and U format the selected text.
 * Each `input` event carries its one contiguous edit onto the runs
 * (`editSegmentsText`), so the browser owns the caret while typing and the
 * runs are never guessed back from flat text.
 *
 * Leaving the cell is reported once through `onEnd`, with the committed
 * segments (or `undefined` when nothing changed) and why: blur, a move
 * (Tab, Shift+Tab, an arrow at the text's edge), Escape, or a paste of
 * tab-separated text that fills cells. `resolveTableCellEditEnd` turns that
 * into the document update and the next cell for every binding.
 *
 * @module render/table-cell-text-editor
 */
import type { TextSegment, TextStyle } from 'ooxml-core/pptx';
import {
	editSegmentsText,
	segmentsPlainText,
} from 'ooxml-core/pptx/editor/render/segments-text-edit';

import type { TableCellCss } from './table-style';
import type { TableCellCaret, TableCellNavigationKey } from './table-cell-navigation';
import { mapTableCellEditorKey } from './table-cell-navigation';
import {
	readCellEditorCaret,
	readCellEditorSelection,
	readCellEditorText,
	renderCellEditorSegments,
	setCellEditorSelection,
} from './table-cell-text-editor-dom';
import { styleSegmentsRange, toLineBreakAt, wordRangeAt } from './table-cell-text-range';
import { isTabularPaste, parseTabularText } from './table-tsv-paste';

/** Why the cell editor closed. */
export type TableCellEditReason =
	| { reason: 'blur' | 'exit' | 'detach' }
	| { reason: 'navigate'; key: TableCellNavigationKey }
	| { reason: 'paste'; grid: string[][] };

/** Why the cell editor closed, with what it committed. */
export type TableCellEditEnd = {
	/** The edited segments, or `undefined` when the cell is unchanged. */
	segments: TextSegment[] | undefined;
} & TableCellEditReason;

export interface TableCellTextEditorOptions {
	/** Where the editor mounts; its content is left to the binding. */
	host: HTMLElement;
	/** The cell's content, from `tableCellSegments(cell)`. */
	segments: TextSegment[];
	/** The cell's resolved style, so runs render as they do on the slide. */
	cellCss?: TableCellCss;
	/** Where the caret starts: select everything (the default), start or end. */
	caret?: TableCellCaret;
	rtl?: boolean;
	spellCheck?: boolean;
	onEnd: (end: TableCellEditEnd) => void;
}

export interface TableCellTextEditorSession {
	/** The contenteditable surface. */
	readonly element: HTMLElement;
	/** The segments as edited so far. */
	segments(): TextSegment[];
	/** Format the selected text (the word at the caret when nothing is selected). */
	applyTextStyle(updates: Partial<TextStyle>): void;
	/** Close the editor; reports `detach` when it was still open. */
	destroy(): void;
}

const TOGGLES = ['bold', 'italic', 'underline'] as const;

/** Open the rich cell editor inside `options.host`. */
export function openTableCellTextEditor(
	options: TableCellTextEditorOptions,
): TableCellTextEditorSession {
	const { host, cellCss } = options;
	const doc = host.ownerDocument;
	const root = doc.createElement('div');
	root.contentEditable = 'true';
	root.setAttribute('data-inline-editor', '');
	root.setAttribute('data-table-cell-editor', '');
	root.setAttribute('role', 'textbox');
	root.setAttribute('aria-multiline', 'true');
	root.spellcheck = options.spellCheck ?? true;
	root.dir = options.rtl ? 'rtl' : 'ltr';
	Object.assign(root.style, {
		outline: 'none',
		minHeight: '1em',
		whiteSpace: 'pre-wrap',
		overflowWrap: 'break-word',
		cursor: 'text',
		width: '100%',
	});
	const initial = options.segments.map((segment) => ({ ...segment }));
	let segments = options.segments.map((segment) => ({ ...segment }));
	let ended = false;

	const render = (start: number, end = start): void => {
		renderCellEditorSegments(root, segments, cellCss);
		setCellEditorSelection(root, start, end);
	};
	const sync = (): void => {
		const text = readCellEditorText(root);
		if (text !== segmentsPlainText(segments)) {
			const caret = readCellEditorSelection(root)?.end;
			segments = editSegmentsText(segments, text, caret);
		}
	};
	const finish = (end: TableCellEditReason): void => {
		if (ended) {
			return;
		}
		sync();
		ended = true;
		const changed = JSON.stringify(segments) !== JSON.stringify(initial);
		options.onEnd({ ...end, segments: changed ? segments : undefined } as TableCellEditEnd);
	};
	const replaceSelection = (insert: string, asLineBreak = false): void => {
		sync();
		const text = segmentsPlainText(segments);
		const { start, end } = readCellEditorSelection(root) ?? {
			start: text.length,
			end: text.length,
		};
		const next = text.slice(0, start) + insert + text.slice(end);
		segments = editSegmentsText(segments, next, start + insert.length);
		if (asLineBreak) {
			segments = toLineBreakAt(segments, start);
		}
		render(start + insert.length);
	};
	const applyTextStyle = (updates: Partial<TextStyle>): void => {
		sync();
		const text = segmentsPlainText(segments);
		let range = readCellEditorSelection(root) ?? { start: 0, end: text.length };
		if (range.start === range.end) {
			range = wordRangeAt(text, range.start);
		}
		if (range.start === range.end) {
			return;
		}
		segments = styleSegmentsRange(segments, range.start, range.end, updates);
		render(range.start, range.end);
	};
	const toggle = (property: (typeof TOGGLES)[number]): void => {
		sync();
		const text = segmentsPlainText(segments);
		let range = readCellEditorSelection(root) ?? { start: 0, end: text.length };
		if (range.start === range.end) {
			range = wordRangeAt(text, range.start);
		}
		let offset = 0;
		let allOn = true;
		for (const segment of segments) {
			const length = segment.isParagraphBreak || segment.isLineBreak ? 1 : segment.text.length;
			const overlaps = offset < range.end && offset + length > range.start;
			if (overlaps && length > 0 && !segment.isParagraphBreak && !segment.isLineBreak) {
				allOn &&= Boolean(segment.style?.[property] ?? cellCssFlag(cellCss, property));
			}
			offset += length;
		}
		applyTextStyle({ [property]: !allOn });
	};

	const onKeyDown = (event: KeyboardEvent): void => {
		event.stopPropagation();
		const action = mapTableCellEditorKey(event, {
			...readCellEditorCaret(root),
			rtl: options.rtl === true,
		});
		if (!action) {
			return;
		}
		event.preventDefault();
		switch (action.kind) {
			case 'navigate':
				finish({ reason: 'navigate', key: action.key });
				return;
			case 'exit':
				finish({ reason: 'exit' });
				return;
			case 'paragraph':
				replaceSelection('\n');
				return;
			case 'line-break':
				replaceSelection('\n', true);
				return;
			case 'tab-character':
				replaceSelection('\t');
				return;
			case 'format':
				toggle(action.property);
				return;
			case 'suppress':
				return;
		}
	};
	const onPaste = (event: ClipboardEvent): void => {
		event.stopPropagation();
		const text = event.clipboardData?.getData('text/plain') ?? '';
		event.preventDefault();
		if (!text) {
			return;
		}
		const grid = isTabularPaste(text) ? parseTabularText(text) : undefined;
		if (grid && (grid.length > 1 || (grid[0]?.length ?? 0) > 1)) {
			finish({ reason: 'paste', grid });
			return;
		}
		replaceSelection((grid?.[0]?.[0] ?? text).replace(/\r\n?/g, '\n'));
	};
	const stop = (event: Event): void => event.stopPropagation();
	const onInput = (): void => sync();
	const onBlur = (): void => finish({ reason: 'blur' });

	root.addEventListener('keydown', onKeyDown);
	root.addEventListener('paste', onPaste);
	root.addEventListener('input', onInput);
	root.addEventListener('blur', onBlur);
	for (const type of ['pointerdown', 'mousedown', 'click', 'dblclick', 'keyup', 'keypress']) {
		root.addEventListener(type, stop);
	}
	host.appendChild(root);
	renderCellEditorSegments(root, segments, cellCss);
	root.focus({ preventScroll: true });
	const length = segmentsPlainText(segments).length;
	const caret = options.caret ?? 'all';
	setCellEditorSelection(root, caret === 'end' ? length : 0, caret === 'start' ? 0 : length);

	return {
		element: root,
		segments: () => {
			if (!ended) {
				sync();
			}
			return segments;
		},
		applyTextStyle,
		destroy: () => {
			finish({ reason: 'detach' });
			root.removeEventListener('blur', onBlur);
			root.remove();
		},
	};
}

function cellCssFlag(css: TableCellCss | undefined, property: (typeof TOGGLES)[number]): boolean {
	if (!css) {
		return false;
	}
	if (property === 'bold') {
		const weight = String(css.fontWeight ?? '');
		return weight === 'bold' || Number(weight) >= 600;
	}
	if (property === 'italic') {
		return css.fontStyle === 'italic';
	}
	return String(css.textDecoration ?? css.textDecorationLine ?? '').includes('underline');
}
