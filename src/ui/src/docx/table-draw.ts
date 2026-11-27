import type { EditorCore } from './editor-core';
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
	const finish = () => {
		controller.abort();
		delete canvas.dataset.tableDraw;
		drawing.delete(canvas);
		core.refreshControls();
	};
	if (mode !== 'paint') {
		void import('./table-freeform')
			.then(({ startFreeformTable }) => {
				if (!signal.aborted) startFreeformTable(core, canvas, mode, signal);
			})
			.catch((error: unknown) => core.host.reportError(error));
		listenForEscape(canvas, signal, finish);
		core.refreshControls();
		return;
	}
	let start: Element | null = null;
	const cellOf = (event: PointerEvent) =>
		event.target instanceof Element ? event.target.closest('td, .dve-print-cell') : null;
	canvas.addEventListener(
		'pointerdown',
		(event) => {
			if (event.button !== 0) return;
			const cell = cellOf(event);
			if (!cell || tableTabReason(core, 'draw', tableTabContext(view.state))) return finish();
			start = cell;
			event.preventDefault();
			event.stopImmediatePropagation();
		},
		{ capture: true, signal },
	);
	canvas.addEventListener(
		'pointercancel',
		() => {
			start = null;
		},
		{ signal },
	);
	canvas.addEventListener(
		'pointerup',
		(event) => {
			const anchor = start;
			start = null;
			const cell = cellOf(event);
			const context = tableTabContext(view.state);
			if (!anchor || cell !== anchor || !context || tableTabReason(core, 'draw', context)) return;
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
			const nearest = [
				['top', Math.abs(rect.top - event.clientY)],
				['bottom', Math.abs(rect.bottom - event.clientY)],
				['left', Math.abs(rect.left - event.clientX)],
				['right', Math.abs(rect.right - event.clientX)],
			] as [CellSide, number][];
			nearest.sort((a, b) => a[1] - b[1]);
			const side = nearest[0]?.[0];
			if (side) paintCellBorder(view, context.pos, target.row, target.column, side, pen);
		},
		{ capture: true, signal },
	);
	listenForEscape(canvas, signal, finish);
	core.refreshControls();
}

/** Esc leaves the drawing tool, as in Word. */
function listenForEscape(canvas: HTMLElement, signal: AbortSignal, finish: () => void): void {
	canvas.getRootNode().addEventListener(
		'keydown',
		(event) => {
			if (!(event instanceof KeyboardEvent) || event.key !== 'Escape') return;
			event.preventDefault();
			finish();
		},
		{ capture: true, signal },
	);
}
