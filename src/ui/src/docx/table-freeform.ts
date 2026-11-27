// Table Layout > Draw Table and Eraser: the pointer tools. A drag in the body draws a one-cell
// table; a drag inside a cell draws a line that splits it; the Eraser merges the two cells beside
// the line it touches (ooxml-core/docx table-draw has the grid rules). Each edit replaces the
// table through the merge-aware table path in one undo step. Loaded when either tool is chosen.
import { DEFAULT_TABLE_BORDERS, type Table } from 'ooxml-core/docx';
import {
	DEFAULT_TABLE_LOOK,
	RESTRUCTURE_TABLE_META,
	drawTableLine,
	drawnTableBox,
	eraseTableEdge,
	type ErasedEdge,
} from 'ooxml-core/docx/ui';
import { twipsFromPixels } from 'ooxml-core/units';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import { Selection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import type { EditorCore } from './editor-core';
import { translate, type LocalizationKey } from './localization';
import { schema } from './schema';
import { paintCellBorder, type CellSide } from './table-border-commands';
import { paragraphIdFactory } from './table-commands';
import { dispatchTableEdit } from './table-edit-transaction';
import { nodeFromResult, skeletonOf, slotsOf } from './table-restructure';
import { freeformTableReason } from './table-tab-context';

export type FreeformMode = 'draw' | 'erase';

/** Drags shorter than this (CSS pixels) are clicks: they draw nothing. */
const MIN_DRAG_PX = 6;

/** A table cell located in the document. */
export interface CellTarget {
	table: ProseMirrorNode;
	tablePos: number;
	nested: boolean;
	/** The cell in the core skeleton: its row and position in that row. */
	row: number;
	index: number;
	/** The cell's position in the table node (its row node and index), for border edits. */
	nodeRow: number;
	nodeColumn: number;
}

interface Hit extends CellTarget {
	element: HTMLElement;
	rtl: boolean;
}

/** The table cell around document position `inside`. */
export function cellTarget(doc: ProseMirrorNode, inside: number): CellTarget | undefined {
	const $pos = doc.resolve(inside);
	for (let depth = $pos.depth; depth >= 2; depth--) {
		const cell = $pos.node(depth);
		if (cell.type.name !== 'tableCell') continue;
		const table = $pos.node(depth - 2);
		if (table.type.name !== 'table') return undefined;
		const slots = slotsOf(table);
		for (let row = 0; row < slots.length; row++) {
			const index = slots[row]?.findIndex((slot) => slot.cell === cell && !slot.offset) ?? -1;
			if (index < 0) continue;
			return {
				table,
				tablePos: $pos.before(depth - 2),
				nested: depth - 3 >= 0 && $pos.node(depth - 3).type.name === 'tableCell',
				row,
				index,
				nodeRow: $pos.index(depth - 2),
				nodeColumn: $pos.index(depth - 1),
			};
		}
		return undefined;
	}
	return undefined;
}

/** The table cell under `target`, in Draft (`td`) or Print Layout (`.dve-print-cell`). */
function hitCell(view: EditorView, target: EventTarget | null): Hit | undefined {
	const element =
		target instanceof Element ? target.closest<HTMLElement>('td, .dve-print-cell') : null;
	if (!element) return undefined;
	let inside: number | undefined;
	if (element instanceof HTMLTableCellElement && view.dom.contains(element)) {
		try {
			inside = view.posAtDOM(element, 0);
		} catch {
			return undefined;
		}
	} else {
		const id = element.dataset.tableCellParagraph;
		if (!id) return undefined;
		view.state.doc.descendants((node, pos) => {
			if (inside !== undefined) return false;
			if (node.type.name === 'paragraph' && node.attrs.id === id) inside = pos + 1;
			return inside === undefined;
		});
	}
	const found = inside === undefined ? undefined : cellTarget(view.state.doc, inside);
	return found && { ...found, element, rtl: getComputedStyle(element).direction === 'rtl' };
}

function replaceTable(view: EditorView, hit: CellTarget, result: Table): boolean {
	const replacement = nodeFromResult(hit.table, result, hit.nested);
	if (!replacement) return false;
	const tr = view.state.tr.replaceWith(
		hit.tablePos,
		hit.tablePos + hit.table.nodeSize,
		replacement,
	);
	tr.setMeta(RESTRUCTURE_TABLE_META, true);
	tr.setSelection(Selection.near(tr.doc.resolve(hit.tablePos + 1), 1));
	dispatchTableEdit(view, tr);
	return true;
}

/**
 * Draws a line across a cell: `fraction` of the way across it (a vertical line) or down it (a
 * horizontal one), in logical order. False when the line cannot split the cell there.
 */
export function drawLineInCell(
	view: EditorView,
	target: CellTarget,
	orientation: 'vertical' | 'horizontal',
	fraction: number,
	defaultWidthTwips?: number,
): boolean {
	const result = drawTableLine(
		skeletonOf(target.table),
		{ orientation, row: target.row, index: target.index, fraction },
		paragraphIdFactory(view),
		defaultWidthTwips ? { defaultWidthTwips } : {},
	);
	return Boolean(result && replaceTable(view, target, result));
}

/**
 * Erases a side of a cell: merges the two cells beside an inner line, or removes the border of an
 * outer one (in a table without merged cells). Returns the reason when nothing changed.
 */
export function eraseCellSide(
	view: EditorView,
	target: CellTarget,
	side: CellSide,
	fraction = 0.5,
): LocalizationKey | undefined {
	const edge: ErasedEdge = { row: target.row, index: target.index, side, fraction };
	const outcome = eraseTableEdge(skeletonOf(target.table), edge, paragraphIdFactory(view));
	if (outcome.kind === 'merged')
		return replaceTable(view, target, outcome.table) ? undefined : 'tableTab.eraseRefused';
	if (outcome.kind === 'outer')
		return target.table.attrs.structureEditable !== false &&
			paintCellBorder(view, target.tablePos, target.nodeRow, target.nodeColumn, side, null)
			? undefined
			: 'tableTab.eraseOuterRefused';
	return 'tableTab.eraseRefused';
}

/** Inserts the one-cell table a drawn rectangle makes at `pos`; false when it is too small. */
export function insertDrawnTable(
	view: EditorView,
	pos: number,
	widthTwips: number,
	heightTwips: number,
): boolean {
	const table = drawnTable(view, widthTwips, heightTwips);
	if (!table) return false;
	const tr = view.state.tr.insert(pos, table);
	// Word keeps a paragraph after every table; one drawn at the end of the body gets one.
	const after = pos + table.nodeSize;
	if (after >= tr.doc.content.size)
		tr.insert(after, schema.nodes.paragraph.create({ id: paragraphIdFactory(view)() }));
	tr.setSelection(Selection.near(tr.doc.resolve(pos + 1), 1));
	dispatchTableEdit(view, tr.scrollIntoView());
	return true;
}

/** The nearest side of the hit cell to a point, and where along that side it lies. */
function nearestSide(hit: Hit, x: number, y: number): { edge: ErasedEdge; distance: number } {
	const rect = hit.element.getBoundingClientRect();
	const across = rect.width ? (x - rect.left) / rect.width : 0.5;
	const down = rect.height ? (y - rect.top) / rect.height : 0.5;
	const sides: [CellSide, number][] = [
		['top', Math.abs(y - rect.top)],
		['bottom', Math.abs(rect.bottom - y)],
		['left', Math.abs(x - rect.left)],
		['right', Math.abs(rect.right - x)],
	];
	sides.sort((a, b) => a[1] - b[1]);
	const [screen, distance] = sides[0] ?? ['top', Infinity];
	// A right-to-left table is mirrored on screen: its logical start is on the right.
	const side: CellSide =
		hit.rtl && (screen === 'left' || screen === 'right')
			? screen === 'left'
				? 'right'
				: 'left'
			: screen;
	const fraction = side === 'left' || side === 'right' ? down : hit.rtl ? 1 - across : across;
	return { edge: { row: hit.row, index: hit.index, side, fraction }, distance };
}

/** Where a new table drawn from the point (x, y) goes: before or after the body block there. */
function insertionPos(view: EditorView, x: number, y: number): number {
	const bounds = view.dom.getBoundingClientRect();
	const found = view.posAtCoords({
		left: Math.min(bounds.right - 1, Math.max(bounds.left + 1, x)),
		top: Math.min(bounds.bottom - 1, Math.max(bounds.top + 1, y)),
	});
	const doc = view.state.doc;
	const $pos = doc.resolve(found?.pos ?? view.state.selection.from);
	if ($pos.depth < 1) return $pos.pos;
	const before = $pos.before(1);
	const block = view.nodeDOM(before);
	const box = block instanceof HTMLElement ? block.getBoundingClientRect() : undefined;
	return box && y > box.top + box.height / 2 ? $pos.after(1) : before;
}

function drawnTable(view: EditorView, width: number, height: number): ProseMirrorNode | undefined {
	const box = drawnTableBox(width, height);
	if (!box) return undefined;
	const newId = paragraphIdFactory(view);
	const properties = { heightTwips: box.heightTwips, heightRule: 'atLeast' };
	const cell = schema.nodes.tableCell.create(
		{ widthTwips: box.widthTwips },
		schema.nodes.paragraph.create({ id: newId() }),
	);
	return schema.nodes.table.create(
		{
			id: `${newId()}-table`,
			borders: JSON.stringify(DEFAULT_TABLE_BORDERS),
			tablePropertiesKnown: true,
			tableLook: DEFAULT_TABLE_LOOK,
			tableLayout: 'fixed',
			grid: JSON.stringify([box.widthTwips]),
			widthTwips: box.widthTwips,
		},
		schema.nodes.tableRow.create({ heightPx: box.heightTwips / 15, properties }, cell),
	);
}

/** Attaches the Draw Table or Eraser pointer tool to the canvas until `signal` aborts. */
export function startFreeformTable(
	core: EditorCore,
	canvas: HTMLElement,
	mode: FreeformMode,
	signal: AbortSignal,
): void {
	const preview = document.createElement('div');
	preview.className = 'dve-table-draw-preview';
	preview.hidden = true;
	const status = document.createElement('div');
	status.className = 'dve-table-draw-status';
	status.setAttribute('role', 'status');
	canvas.append(preview, status);
	signal.addEventListener('abort', () => {
		preview.remove();
		status.remove();
	});
	const say = (key: LocalizationKey | undefined) => {
		status.textContent = key ? translate(core.locale, key) : '';
	};
	let start: { x: number; y: number; hit: Hit | undefined } | undefined;

	const showPreview = (x: number, y: number) => {
		if (!start) return;
		let left = Math.min(start.x, x);
		let top = Math.min(start.y, y);
		let width = Math.abs(x - start.x);
		let height = Math.abs(y - start.y);
		if (mode === 'draw' && start.hit) {
			// Inside a cell the pen draws a straight line along the drag's longer direction.
			const cell = start.hit.element.getBoundingClientRect();
			if (width >= height) [top, height, left, width] = [start.y, 0, cell.left, cell.width];
			else [left, width, top, height] = [start.x, 0, cell.top, cell.height];
		}
		Object.assign(preview.style, {
			left: `${left}px`,
			top: `${top}px`,
			width: `${width}px`,
			height: `${height}px`,
		});
		preview.hidden = mode === 'erase' || Math.max(width, height) < MIN_DRAG_PX;
	};

	const view = () => core.view;
	canvas.addEventListener(
		'pointerdown',
		(event) => {
			const editor = view();
			if (event.button !== 0 || !editor || !(event.target instanceof Node)) return;
			if (!canvas.contains(event.target) || preview.contains(event.target)) return;
			event.preventDefault();
			event.stopImmediatePropagation();
			const reason = freeformTableReason(core);
			say(reason);
			if (reason) return;
			start = { x: event.clientX, y: event.clientY, hit: hitCell(editor, event.target) };
			try {
				canvas.setPointerCapture(event.pointerId);
			} catch {
				// Synthetic pointers cannot be captured; the drag still ends on pointerup.
			}
		},
		{ capture: true, signal },
	);
	canvas.addEventListener('pointermove', (event) => showPreview(event.clientX, event.clientY), {
		signal,
	});
	canvas.addEventListener(
		'pointercancel',
		() => {
			start = undefined;
			preview.hidden = true;
		},
		{ signal },
	);
	canvas.addEventListener(
		'pointerup',
		(event) => {
			const from = start;
			start = undefined;
			preview.hidden = true;
			const editor = view();
			if (!from || !editor) return;
			event.preventDefault();
			event.stopImmediatePropagation();
			if (freeformTableReason(core)) return;
			const dx = event.clientX - from.x;
			const dy = event.clientY - from.y;
			const dragged = Math.max(Math.abs(dx), Math.abs(dy)) >= MIN_DRAG_PX;
			if (mode === 'erase') return erase(editor, from, event);
			if (!dragged) return;
			if (from.hit) return split(editor, from.hit, from, event, Math.abs(dx) >= Math.abs(dy));
			insertDrawn(editor, from, event);
		},
		{ capture: true, signal },
	);

	const split = (
		editor: EditorView,
		hit: Hit,
		from: { x: number; y: number },
		end: PointerEvent,
		horizontal: boolean,
	) => {
		const rect = hit.element.getBoundingClientRect();
		const fraction = horizontal
			? ((from.y + end.clientY) / 2 - rect.top) / (rect.height || 1)
			: ((from.x + end.clientX) / 2 - rect.left) / (rect.width || 1);
		const drawn = drawLineInCell(
			editor,
			hit,
			horizontal ? 'horizontal' : 'vertical',
			!horizontal && hit.rtl ? 1 - fraction : fraction,
			twipsFromPixels(core.contentWidth()),
		);
		say(drawn ? undefined : 'tableTab.drawLineRefused');
	};

	const erase = (
		editor: EditorView,
		from: { x: number; y: number; hit: Hit | undefined },
		end: PointerEvent,
	) => {
		// A drag erases the line nearest either end of it.
		const last = hitCell(editor, document.elementFromPoint(end.clientX, end.clientY));
		const candidates = [
			from.hit && { hit: from.hit, ...nearestSide(from.hit, from.x, from.y) },
			last && { hit: last, ...nearestSide(last, end.clientX, end.clientY) },
		].filter((entry) => entry !== undefined);
		candidates.sort((a, b) => a.distance - b.distance);
		const choice = candidates[0];
		if (!choice) return;
		say(eraseCellSide(editor, choice.hit, choice.edge.side, choice.edge.fraction));
	};

	const insertDrawn = (editor: EditorView, from: { x: number; y: number }, end: PointerEvent) => {
		// Undo the canvas zoom so the table has the size drawn at 100 percent.
		const bounds = editor.dom.getBoundingClientRect();
		const scale = editor.dom.offsetWidth ? bounds.width / editor.dom.offsetWidth : 1;
		const width = twipsFromPixels(Math.abs(end.clientX - from.x) / (scale || 1));
		const height = twipsFromPixels(Math.abs(end.clientY - from.y) / (scale || 1));
		const maxWidth = twipsFromPixels(core.contentWidth());
		const pos = insertionPos(editor, Math.min(from.x, end.clientX), Math.min(from.y, end.clientY));
		say(
			insertDrawnTable(editor, pos, Math.min(width, maxWidth || width), height)
				? undefined
				: 'tableTab.drawTooSmall',
		);
	};
}
