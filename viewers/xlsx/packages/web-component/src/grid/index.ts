// `mountGrid(ctx, container)`: the virtualized worksheet grid. Wires the view, selection,
// in-cell editor, keyboard, pointer, clipboard, menus, drawings, indicators and accessibility,
// registers the grid's `edit.*` commands and attaches the GridController to the context.
import type { CellRange } from '@christophervr/xlsx-core';
import { createGridClipboard } from '../clipboard.js';
import type { EditorContext, GridController, Selection } from '../context.js';
import { CellEditor } from './cell-editor.js';
import { ensureStyle } from './dom.js';
import { drawingKeyDown } from './drawing-keys.js';
import { DrawingLayer } from './drawings.js';
import { GridA11y } from './grid-a11y.js';
import { gridCommands } from './grid-commands.js';
import { editorKeyDown, readyKeyDown, type KeyHost } from './grid-keys.js';
import { openMenuAtActive, wireMenus } from './grid-menus.js';
import { wirePointer } from './grid-pointer.js';
import { GridSelection } from './grid-selection.js';
import { GridView } from './grid-view.js';
import { followLink, wireTooltips } from './indicators.js';
import { selectCell } from './selection-ops.js';
import { GRID_CSS } from './styles.js';
import { ValidationList } from './validation-list.js';

export { editBridge, type EditBridge, type EditMode, type EditState } from './edit-bridge.js';

/** The GridController plus the grid's status-bar mode. */
export interface GridControllerExt extends GridController {
	/** Excel's mode indicator: 'ready' | 'enter' | 'edit' | 'point'. */
	mode(): string;
}

const structuralKinds = new Set(['structure', 'sheets', 'view', 'undo', 'redo', 'batch']);

export function mountGrid(ctx: EditorContext, container: HTMLElement): () => void {
	ensureStyle(ctx.root, 'xg', GRID_CSS);
	const view = new GridView(ctx, container);
	const selection = new GridSelection(view);
	const editor = new CellEditor(view, (move) => {
		if (move !== 'none') selection.advance(move);
	});
	const sink = editor.field.input;
	const drawings = new DrawingLayer(view);
	const list = new ValidationList(view, () => editor.focus());
	const a11y = new GridA11y(view, sink);
	let marqueeSheet = -1;
	const clipboard = createGridClipboard({
		ctx,
		sink,
		isEditing: () => editor.editing,
		setMarquee(range, sheet) {
			marqueeSheet = sheet;
			view.overlay.copyRange = range;
			view.schedule();
		},
		selectRange(range: CellRange) {
			selection.set({
				sheet: view.sheetIndex(),
				active: { ...range.start },
				anchor: { ...range.start },
				ranges: [range],
			});
		},
	});
	const keyHost: KeyHost = {
		view,
		selection,
		editor,
		openCellMenu: () => openMenuAtActive(view, () => editor.focus()),
		openValidationList: () => list.open(true),
		clearMarquee: () => clipboard.clearMarquee(),
		prepareCopy: (cut) => clipboard.prepareNativeCopy(cut),
	};
	const onKey = (event: KeyboardEvent) => {
		if (editor.editing) return editorKeyDown(keyHost, event);
		if (drawingKeyDown(ctx, drawings, event)) return;
		readyKeyDown(keyHost, event);
	};
	sink.addEventListener('keydown', onKey);

	const disposers: (() => void)[] = [
		wirePointer({
			view,
			selection,
			editor,
			followLink: (row, col) => followLink(view, selection, row, col),
			fill(source, target) {
				ctx.session()?.fill(view.sheetIndex(), source, target);
				selection.set({ ...selection.get(), ranges: [target] });
			},
		}),
		wireTooltips(view),
		wireMenus(view, () => editor.focus()),
	];
	view.corner.addEventListener('pointerdown', (event) => {
		event.preventDefault();
		event.stopPropagation();
		if (editor.editing && !editor.commit('none', false)) return;
		selection.all();
		editor.focus();
	});
	view.view.addEventListener('pointerdown', (event) => {
		if (!(event.target as Element).closest('.xg-obj')) drawings.deselect();
	});

	ctx.commands.registerAll(
		gridCommands({
			clipboard,
			cancelEdit: () => editor.cancel(),
			isEditing: () => editor.editing,
			selectAll: () => selection.all(),
			openValidationList: () => list.open(true),
		}),
	);

	// The editor core keeps each sheet's selection (and emits selection-change, once, with the
	// shared A1 formatter); the grid only falls back to the file's selection when it mounts on a
	// sheet the selection does not belong to.
	let shownSheet = view.sheet();
	let shownWorkbook = view.workbook();
	const restoreSelection = () => {
		const sheet = view.sheet();
		const index = view.sheetIndex();
		const fromFile = sheet?.view.selection;
		if (fromFile?.ranges.length)
			ctx.selection.set({
				sheet: index,
				active: fromFile.active,
				anchor: fromFile.active,
				ranges: fromFile.ranges,
			});
		else ctx.selection.set(selectCell(index, sheet, { row: 0, col: 0 }));
	};
	const onSelection = (next: Selection) => {
		a11y.update(next, editor.editing);
		view.schedule();
		ctx.requestRender();
	};
	const offSelection = ctx.selection.onChange(onSelection);
	const onModel = (change: unknown) => {
		const workbook = view.workbook();
		const sheet = view.sheet();
		drawings.invalidate();
		if (workbook !== shownWorkbook || sheet !== shownSheet) {
			if (editor.editing) editor.end();
			if (marqueeSheet >= 0) clipboard.clearMarquee();
			shownWorkbook = workbook;
			shownSheet = sheet;
			view.scroller.scrollLeft = 0;
			view.scroller.scrollTop = 0;
			view.rebuild();
			if (ctx.selection.get().sheet !== view.sheetIndex() || !workbook) restoreSelection();
			const top = sheet?.view.topLeft;
			if (top) view.reveal(top.row, top.col);
			return;
		}
		const kind =
			typeof change === 'object' && change && 'kind' in change
				? String((change as { kind: unknown }).kind)
				: '';
		const structural =
			typeof change === 'object' &&
			change !== null &&
			(change as { structural?: unknown }).structural === true;
		if (structural || structuralKinds.has(kind) || !kind) view.rebuild();
		else view.invalidate();
	};
	const offModel = ctx.onModelChange(onModel);

	const resize =
		typeof ResizeObserver === 'undefined' ? undefined : new ResizeObserver(() => view.schedule());
	resize?.observe(view.root);

	const controller: GridControllerExt = {
		focus: () => editor.focus(),
		scrollTo: (address) => view.reveal(address.row, address.col),
		invalidate: () => view.invalidate(),
		beginEdit(initialText) {
			if (initialText === undefined) editor.begin({ mode: 'edit' });
			else editor.begin({ mode: 'enter', text: initialText });
		},
		commitEdit: () => editor.commit('none', false),
		cancelEdit: () => editor.cancel(),
		isEditing: () => editor.editing,
		measureText: (text, font) => view.measurer.measure(text, font),
		zoom: () => view.zoom(),
		setZoom: (percent) => view.setZoom(percent),
		mode: () => editor.bridge.state().mode,
	};
	ctx.attachGrid(controller);
	if (ctx.selection.get().sheet !== view.sheetIndex()) restoreSelection();
	a11y.update(ctx.selection.get(), false);

	return () => {
		offSelection();
		offModel();
		resize?.disconnect();
		sink.removeEventListener('keydown', onKey);
		for (const dispose of disposers) dispose();
		clipboard.destroy();
		list.close();
		drawings.destroy();
		editor.destroy();
		a11y.destroy();
		view.destroy();
	};
}
