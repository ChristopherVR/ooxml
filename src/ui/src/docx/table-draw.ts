import { Selection } from 'prosemirror-state';
import type { EditorCore } from './editor-core';
import { executeTableCommand } from './table-commands';
import { paintCellBorder, type BorderPen, type CellSide } from './table-border-commands';
import { tableCells, tableTabContext, tableTabReason } from './table-tab-context';
import type { TableTabAction } from './table-ribbon';

type Mode = Extract<TableTabAction, { key: 'draw' }>['value'];
const drawing = new WeakMap<HTMLElement, AbortController>();

export function toggleTableDraw(core: EditorCore, mode: Mode, pen: BorderPen): void {
	const canvas = core.shell.canvas;
	const view = core.view;
	if (!canvas || !view) return;
	const active = canvas.dataset.tableDraw === mode;
	drawing.get(canvas)?.abort();
	delete canvas.dataset.tableDraw;
	if (active) {
		core.refreshControls();
		return;
	}
	const controller = new AbortController();
	drawing.set(canvas, controller);
	canvas.dataset.tableDraw = mode;
	const { signal } = controller;
	let start: { cell: Element; x: number; y: number } | undefined;
	const finish = () => {
		controller.abort();
		delete canvas.dataset.tableDraw;
		drawing.delete(canvas);
		core.refreshControls();
	};
	const cellOf = (event: PointerEvent) =>
		event.target instanceof Element ? event.target.closest('td, .dve-print-cell') : null;
	canvas.addEventListener(
		'pointerdown',
		(event) => {
			if (event.button !== 0) return;
			const cell = cellOf(event);
			if (!cell) return finish();
			start = { cell, x: event.clientX, y: event.clientY };
			event.preventDefault();
			event.stopImmediatePropagation();
		},
		{ capture: true, signal },
	);
	canvas.addEventListener(
		'pointercancel',
		() => {
			start = undefined;
		},
		{ signal },
	);
	canvas.addEventListener(
		'pointerup',
		(event) => {
			const anchor = start;
			start = undefined;
			const cell = cellOf(event);
			const context = tableTabContext(view.state);
			if (!anchor || cell !== anchor.cell || !context || tableTabReason(core, 'draw', context))
				return;
			const cells = tableCells(context.node, context.pos);
			const target =
				cell instanceof HTMLTableCellElement
					? cells.find(({ pos }) => view.nodeDOM(pos) === cell)
					: cells.find(
							({ node }) =>
								cell instanceof HTMLElement &&
								node.firstChild?.attrs.id === cell.dataset.tableCellParagraph,
						);
			if (!target) return;
			event.preventDefault();
			event.stopImmediatePropagation();
			const rect = cell.getBoundingClientRect();
			const dx = Math.abs(event.clientX - anchor.x);
			const dy = Math.abs(event.clientY - anchor.y);
			if (mode === 'draw' && (dx > rect.width / 2 || dy > rect.height / 2)) {
				const horizontal = dx >= dy;
				const center = horizontal ? (event.clientY + anchor.y) / 2 : (event.clientX + anchor.x) / 2;
				const offset = horizontal
					? (center - rect.top) / rect.height
					: (center - rect.left) / rect.width;
				if (offset >= 0.2 && offset <= 0.8) {
					view.dispatch(
						view.state.tr
							.setSelection(Selection.near(view.state.doc.resolve(target.pos + 1)))
							.setMeta('addToHistory', false),
					);
					executeTableCommand(view, 'splitCell', {
						rows: horizontal ? 2 : 1,
						columns: horizontal ? 1 : 2,
					});
					return finish();
				}
			}
			const nearest = [
				['top', Math.abs(rect.top - event.clientY)],
				['bottom', Math.abs(rect.bottom - event.clientY)],
				['left', Math.abs(rect.left - event.clientX)],
				['right', Math.abs(rect.right - event.clientX)],
			] as [CellSide, number][];
			nearest.sort((a, b) => a[1] - b[1]);
			const side = nearest[0]?.[0];
			if (side)
				paintCellBorder(
					view,
					context.pos,
					target.row,
					target.column,
					side,
					mode === 'erase' ? null : pen,
				);
		},
		{ capture: true, signal },
	);
	const cancel = (event: KeyboardEvent) => {
		if (event.key === 'Escape') {
			event.preventDefault();
			finish();
		}
	};
	canvas.getRootNode().addEventListener(
		'keydown',
		(event) => {
			if (event instanceof KeyboardEvent) cancel(event);
		},
		{ capture: true, signal },
	);
	core.refreshControls();
}
