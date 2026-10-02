// Mouse, pen and touch input on the grid: cell / header / corner selection with drag and
// autoscroll, Ctrl multi-range, point mode while editing a formula, fill-handle drags, header
// border resizing and double-click autofit, hyperlink Ctrl+click, and touch tap-to-select.
import { normalizeRange, type CellAddress, type CellRange } from '@christophervr/xlsx-core';
import type { CellEditor } from './cell-editor.js';
import type { Hit } from './geometry.js';
import type { GridSelection } from './grid-selection.js';
import type { GridView } from './grid-view.js';
import { startHeaderResize, autoFit } from './header-resize.js';
import { addRange, extendSelection, selectCell, spanRange } from './selection-ops.js';

export interface PointerHost {
	view: GridView;
	selection: GridSelection;
	editor: CellEditor;
	followLink(row: number, col: number): boolean;
	fill(source: CellRange, target: CellRange): void;
}

type DragKind = 'cells' | 'rows' | 'cols' | 'point' | 'fill';

/** The range a fill-handle drag to `at` produces (extends the source along one axis). */
export function fillTarget(source: CellRange, at: CellAddress): CellRange | undefined {
	const s = normalizeRange(source);
	const down = at.row - s.end.row;
	const up = s.start.row - at.row;
	const right = at.col - s.end.col;
	const left = s.start.col - at.col;
	const vertical = Math.max(down, up);
	const horizontal = Math.max(right, left);
	if (vertical <= 0 && horizontal <= 0) return undefined;
	if (vertical >= horizontal) {
		if (down > 0) return { start: s.start, end: { row: at.row, col: s.end.col } };
		return { start: { row: at.row, col: s.start.col }, end: s.end };
	}
	if (right > 0) return { start: s.start, end: { row: s.end.row, col: at.col } };
	return { start: { row: s.start.row, col: at.col }, end: s.end };
}

export function wirePointer(host: PointerHost): () => void {
	const { view, selection, editor } = host;
	const el = view.view;
	let drag:
		| {
				kind: DragKind;
				anchor: CellAddress;
				pointer: number;
				x: number;
				y: number;
				last?: CellAddress;
		  }
		| undefined;
	let autoscroll: ReturnType<typeof setInterval> | undefined;
	let touch: { x: number; y: number; id: number } | undefined;

	const stopAutoscroll = () => {
		if (autoscroll) clearInterval(autoscroll);
		autoscroll = undefined;
	};

	const clampedHit = (x: number, y: number): Hit => {
		const g = view.geometry;
		const cx = Math.min(Math.max(x, g.headerW + 1), g.width - 2);
		const cy = Math.min(Math.max(y, g.headerH + 1), g.height - 2);
		return { area: 'cell', row: g.rowAtY(cy), col: g.colAtX(cx) };
	};

	const track = (x: number, y: number) => {
		if (!drag) return;
		const sheet = view.sheet();
		const hit = clampedHit(x, y);
		const at = { row: hit.row, col: hit.col };
		if (drag.last && drag.last.row === at.row && drag.last.col === at.col) return;
		drag.last = at;
		const current = selection.get();
		if (drag.kind === 'cells') selection.set(extendSelection(current, sheet, at));
		else if (drag.kind === 'rows') selection.rows(drag.anchor.row, at.row);
		else if (drag.kind === 'cols') selection.cols(drag.anchor.col, at.col);
		else if (drag.kind === 'point') editor.point.pointRange(drag.anchor, at);
		else if (drag.kind === 'fill') {
			const source = current.ranges[current.ranges.length - 1];
			view.overlay.fillPreview = source ? fillTarget(source, at) : undefined;
			view.schedule();
		}
	};

	const edgeScroll = (x: number, y: number) => {
		const g = view.geometry;
		const dx = x < g.headerW + g.frozenW ? -1 : x > g.width ? 1 : 0;
		const dy = y < g.headerH + g.frozenH ? -1 : y > g.height ? 1 : 0;
		const scrollX = drag?.kind !== 'rows' ? dx : 0;
		const scrollY = drag?.kind !== 'cols' ? dy : 0;
		stopAutoscroll();
		if (!scrollX && !scrollY) return;
		autoscroll = setInterval(() => {
			view.scroller.scrollLeft += scrollX * 24;
			view.scroller.scrollTop += scrollY * 20;
			view.paint();
			track(x, y);
		}, 40);
	};

	const onDown = (event: PointerEvent) => {
		const target = event.target as Element;
		if (target.closest('.xg-editor-host, .xg-obj, .xg-dd, .xg-popup, .xg-assist')) return;
		const { x, y } = view.clientToView(event.clientX, event.clientY);
		const g = view.geometry;
		const hit = g.hit(x, y);
		if (event.pointerType === 'touch') {
			touch = { x: event.clientX, y: event.clientY, id: event.pointerId };
			return;
		}
		if (event.button === 2) {
			if (hit.area === 'cell' && !inSelection(selection.get().ranges, hit)) selection.cellAt(hit);
			editor.focus();
			return;
		}
		if (event.button !== 0) return;
		event.preventDefault();
		const sheet = view.sheet();
		const editable = !view.ctx.readOnly();
		if (target.closest('[data-handle="fill"]') && editable) {
			drag = { kind: 'fill', anchor: selection.get().active, pointer: event.pointerId, x, y };
			el.setPointerCapture?.(event.pointerId);
			return;
		}
		if (hit.area === 'col-header' || hit.area === 'row-header') {
			const border = hit.area === 'col-header' ? g.colBorderAt(x) : g.rowBorderAt(y);
			if (border !== undefined && editable && !editor.editing) {
				startHeaderResize(view, hit.area === 'col-header' ? 'col' : 'row', border, event);
				return;
			}
		}
		if (editor.editing) {
			if (hit.area === 'cell' && editor.point.canPoint()) {
				drag = {
					kind: 'point',
					anchor: { row: hit.row, col: hit.col },
					pointer: event.pointerId,
					x,
					y,
				};
				editor.point.pointRange(drag.anchor, drag.anchor);
				el.setPointerCapture?.(event.pointerId);
				return;
			}
			if (!editor.commit('none', false)) return;
		}
		editor.focus();
		const ctrl = event.ctrlKey || event.metaKey;
		if (hit.area === 'corner') return selection.all();
		if (hit.area === 'col-header') {
			const anchorCol = event.shiftKey ? selection.get().anchor.col : hit.col;
			selection.cols(anchorCol, hit.col, ctrl);
			drag = { kind: 'cols', anchor: { row: 0, col: anchorCol }, pointer: event.pointerId, x, y };
		} else if (hit.area === 'row-header') {
			const anchorRow = event.shiftKey ? selection.get().anchor.row : hit.row;
			selection.rows(anchorRow, hit.row, ctrl);
			drag = { kind: 'rows', anchor: { row: anchorRow, col: 0 }, pointer: event.pointerId, x, y };
		} else if (hit.area === 'cell') {
			const at = { row: hit.row, col: hit.col };
			if (ctrl && view.sheet() && host.followLink(at.row, at.col)) return;
			const current = selection.get();
			if (event.shiftKey) selection.set(extendSelection(current, sheet, at));
			else if (ctrl) selection.set(addRange(current, sheet, at));
			else selection.set(selectCell(view.sheetIndex(), sheet, at));
			drag = { kind: 'cells', anchor: at, pointer: event.pointerId, x, y };
		} else return;
		el.setPointerCapture?.(event.pointerId);
	};

	const onMove = (event: PointerEvent) => {
		const { x, y } = view.clientToView(event.clientX, event.clientY);
		if (!drag) {
			const g = view.geometry;
			const hit = g.hit(x, y);
			const editable = !view.ctx.readOnly();
			const colBorder = editable && hit.area === 'col-header' && g.colBorderAt(x) !== undefined;
			const rowBorder = editable && hit.area === 'row-header' && g.rowBorderAt(y) !== undefined;
			el.classList.toggle('xg-col-resize', colBorder);
			el.classList.toggle('xg-row-resize', rowBorder);
			return;
		}
		if (event.pointerId !== drag.pointer) return;
		drag.x = x;
		drag.y = y;
		track(x, y);
		edgeScroll(x, y);
	};

	const onUp = (event: PointerEvent) => {
		if (touch && event.pointerId === touch.id) {
			const moved = Math.hypot(event.clientX - touch.x, event.clientY - touch.y) > 8;
			touch = undefined;
			if (moved) return;
			const hit = view.hitClient(event.clientX, event.clientY);
			if (hit.area !== 'cell') return;
			if (editor.editing && editor.point.pointRange(hit, hit)) return;
			if (editor.editing && !editor.commit('none', false)) return;
			selection.cellAt(hit);
			editor.focus();
			return;
		}
		if (!drag || event.pointerId !== drag.pointer) return;
		stopAutoscroll();
		const finished = drag;
		drag = undefined;
		el.releasePointerCapture?.(event.pointerId);
		if (finished.kind === 'fill') {
			const target = view.overlay.fillPreview;
			view.overlay.fillPreview = undefined;
			const source = selection.get().ranges[selection.get().ranges.length - 1];
			if (target && source) host.fill(source, target);
			view.schedule();
		}
		if (finished.kind === 'point') editor.focus();
	};

	const onDblClick = (event: MouseEvent) => {
		const target = event.target as Element;
		if (target.closest('.xg-editor-host, .xg-obj, .xg-popup')) return;
		const { x, y } = view.clientToView(event.clientX, event.clientY);
		const g = view.geometry;
		const hit = g.hit(x, y);
		if (view.ctx.readOnly()) return;
		if (hit.area === 'col-header') {
			const col = g.colBorderAt(x);
			if (col !== undefined) autoFit(view, 'col', col);
			return;
		}
		if (hit.area === 'row-header') {
			const row = g.rowBorderAt(y);
			if (row !== undefined) autoFit(view, 'row', row);
			return;
		}
		if (hit.area === 'cell' && !editor.editing) editor.begin({ mode: 'edit' });
	};

	el.addEventListener('pointerdown', onDown);
	el.addEventListener('pointermove', onMove);
	el.addEventListener('pointerup', onUp);
	el.addEventListener('pointercancel', onUp);
	el.addEventListener('dblclick', onDblClick);
	return () => {
		stopAutoscroll();
		el.removeEventListener('pointerdown', onDown);
		el.removeEventListener('pointermove', onMove);
		el.removeEventListener('pointerup', onUp);
		el.removeEventListener('pointercancel', onUp);
		el.removeEventListener('dblclick', onDblClick);
	};
}

const inSelection = (ranges: CellRange[], at: CellAddress): boolean =>
	ranges.some(
		(r) =>
			at.row >= r.start.row && at.row <= r.end.row && at.col >= r.start.col && at.col <= r.end.col,
	);

export { spanRange };
