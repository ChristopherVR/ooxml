// @vitest-environment jsdom
import type { TablePptxElement, TextSegment } from 'ooxml-core/pptx';
import { afterEach, describe, expect, it } from 'vitest';

import { resolveTableCellEditEnd, resolveTableCellsPaste } from './table-cell-edit-outcome';
import type { TableCellEditEnd } from './table-cell-text-editor';
import { openTableCellTextEditor } from './table-cell-text-editor';
import { readCellEditorText, setCellEditorSelection } from './table-cell-text-editor-dom';

const runs = (): TextSegment[] => [
	{ text: 'Revenue ', style: {} },
	{ text: 'grew', style: { bold: true } },
	{ text: '\n', isParagraphBreak: true, style: {} },
	{ text: 'Second', style: { italic: true } },
];

function open(segments = runs()) {
	const host = document.createElement('div');
	document.body.appendChild(host);
	const ends: TableCellEditEnd[] = [];
	const session = openTableCellTextEditor({ host, segments, onEnd: (end) => ends.push(end) });
	return { host, session, ends, root: session.element };
}

/** Type over the selection the way the browser does: edit a text node, then `input`. */
function typeAt(root: HTMLElement, spanIndex: number, text: string): void {
	const span = root.querySelectorAll('span')[spanIndex]!;
	span.firstChild!.textContent = text;
	root.dispatchEvent(new Event('input', { bubbles: true }));
}

function press(root: HTMLElement, key: string, init: KeyboardEventInit = {}): void {
	root.dispatchEvent(
		new KeyboardEvent('keydown', { key, bubbles: true, cancelable: true, ...init }),
	);
}

const table = (): TablePptxElement =>
	({
		id: 't',
		type: 'table',
		x: 0,
		y: 0,
		width: 100,
		height: 50,
		tableData: {
			columnWidths: [0.5, 0.5],
			rows: [{ cells: [{ text: 'Revenue grew\nSecond' }, { text: 'b' }] }],
		},
	}) as unknown as TablePptxElement;

afterEach(() => {
	document.body.replaceChildren();
});

describe('openTableCellTextEditor', () => {
	it('renders paragraphs and runs and reads the same text back', () => {
		const { root } = open();
		expect(root.querySelectorAll('[data-cell-paragraph]')).toHaveLength(2);
		expect(readCellEditorText(root)).toBe('Revenue grew\nSecond');
		expect(root.querySelectorAll('span')[1]!.style.fontWeight).toBe('bold');
	});

	it('keeps run formatting when a word is retyped, then moves on Tab', () => {
		const { root, ends } = open();
		setCellEditorSelection(root, 6);
		typeAt(root, 0, 'Income ');
		press(root, 'Tab');
		expect(ends).toHaveLength(1);
		expect(ends[0]).toMatchObject({ reason: 'navigate', key: 'next' });
		const segments = ends[0]!.segments!;
		expect(segments.map((s) => s.text)).toEqual(['Income ', 'grew', '\n', 'Second']);
		expect(segments[1]!.style.bold).toBe(true);
		expect(segments[3]!.style.italic).toBe(true);
	});

	it('reports an unchanged cell as no edit, and Escape as exit', () => {
		const { root, ends } = open();
		press(root, 'Escape');
		expect(ends).toEqual([{ reason: 'exit', segments: undefined }]);
	});

	it('adds a paragraph on Enter at the caret', () => {
		const { root, ends } = open();
		setCellEditorSelection(root, 8);
		press(root, 'Enter');
		expect(readCellEditorText(root)).toBe('Revenue \ngrew\nSecond');
		root.dispatchEvent(new Event('blur'));
		expect(ends[0]!.reason).toBe('blur');
		expect(ends[0]!.segments!.filter((s) => s.isParagraphBreak)).toHaveLength(2);
	});

	it('bolds the selected characters on Ctrl+B', () => {
		const { root, session } = open();
		setCellEditorSelection(root, 0, 3);
		press(root, 'b', { ctrlKey: true });
		const [first, second] = session.segments();
		expect(first).toMatchObject({ text: 'Rev', style: { bold: true } });
		expect(second).toMatchObject({ text: 'enue ' });
		expect(second!.style.bold).toBeUndefined();
	});

	it('turns a pasted tab-separated block into a grid paste', () => {
		const { root, ends } = open();
		const event = new Event('paste', { bubbles: true, cancelable: true }) as ClipboardEvent;
		Object.defineProperty(event, 'clipboardData', {
			value: { getData: () => 'x\ty\r\nz\tw\r\n' },
		});
		root.dispatchEvent(event);
		expect(event.defaultPrevented).toBe(true);
		expect(ends[0]).toMatchObject({
			reason: 'paste',
			grid: [
				['x', 'y'],
				['z', 'w'],
			],
		});
	});

	it('does not report twice when destroyed after it ended', () => {
		const { root, session, ends } = open();
		press(root, 'Escape');
		session.destroy();
		expect(ends).toHaveLength(1);
		expect(root.isConnected).toBe(false);
	});
});

describe('resolveTableCellEditEnd', () => {
	it('commits the segments and appends a row for Tab in the last cell', () => {
		const segments: TextSegment[] = [{ text: 'B', style: { bold: true } }];
		const outcome = resolveTableCellEditEnd(table(), 0, 1, {
			reason: 'navigate',
			key: 'next',
			segments,
		});
		const rows = outcome.element!.tableData!.rows;
		expect(rows).toHaveLength(2);
		expect(rows[0]!.cells[1]!.textRuns).toEqual([{ text: 'B', bold: true }]);
		expect(outcome.edit).toEqual({ row: 1, col: 0, caret: 'all' });
	});

	it('leaves the element untouched for an unchanged blur and selects on exit', () => {
		expect(resolveTableCellEditEnd(table(), 0, 0, { reason: 'blur', segments: undefined })).toEqual(
			{},
		);
		expect(
			resolveTableCellEditEnd(table(), 0, 0, { reason: 'exit', segments: undefined }).select,
		).toEqual({ row: 0, col: 0 });
	});

	it('fills cells from the anchor for a grid paste, growing the table', () => {
		const outcome = resolveTableCellEditEnd(table(), 0, 1, {
			reason: 'paste',
			segments: undefined,
			grid: [['1', '2']],
		});
		expect(outcome.element!.tableData!.rows[0]!.cells.map((c) => c.text)).toEqual([
			'Revenue grew\nSecond',
			'1',
			'2',
		]);
	});

	it('pastes into a selected range from its top-left cell', () => {
		const next = resolveTableCellsPaste(
			table(),
			[
				{ row: 0, col: 1 },
				{ row: 0, col: 0 },
			],
			'p\tq\n',
		);
		expect(next!.tableData!.rows[0]!.cells.map((c) => c.text)).toEqual(['p', 'q']);
	});
});
