import type { Node } from 'prosemirror-model';
import type { EditorState } from 'prosemirror-state';
import { forcesTrackChanges } from 'ooxml-core/docx/ui';
import type { EditorCore } from './editor-core';
import type { TableTabAction } from './table-ribbon';
import { loadTablePanels } from './ribbon';

export function tableTabContext(state: EditorState) {
	const { $from } = state.selection;
	const $to = state.selection.ranges.at(-1)?.$to ?? state.selection.$to;
	for (let depth = $from.depth - 2; depth >= 1; depth--) {
		const node = $from.node(depth);
		if (node.type.name !== 'table') continue;
		if ($to.depth < depth + 2 || $to.node(depth) !== node) return;
		return {
			node,
			pos: $from.before(depth),
			top: $from.index(depth),
			bottom: $to.index(depth),
			firstColumn: $from.index(depth + 1),
			lastColumn: $to.index(depth + 1),
			plain: node.attrs.structureEditable !== false,
		};
	}
}

type Context = NonNullable<ReturnType<typeof tableTabContext>>;
export function canRepeatAt(context: Context): boolean {
	for (let row = 0; row < context.top; row++)
		if (!context.node.child(row).attrs.properties?.header) return false;
	return true;
}

export function tableTabReason(core: EditorCore, key: TableTabAction['key'], context?: Context) {
	if (!context) return 'tableTab.selectCell';
	if (key === 'gridlines') return;
	if (!context.plain) return 'tableTab.rectangular';
	if (key === 'select') return;
	if (
		key !== 'formula' &&
		key !== 'pen' &&
		(core.model.trackChanges || forcesTrackChanges(core.model.protection))
	)
		return 'tableTab.tracking';
	if (
		!core.view?.editable ||
		!core.canEditBody() ||
		core.targetView() !== core.view ||
		core.collab.active ||
		core.model.protection?.enforced
	)
		return 'tableTab.editable';
	if (key === 'repeat' && !canRepeatAt(context)) return 'tableTab.headerStart';
	if (
		key === 'formula' &&
		(context.top !== context.bottom || context.firstColumn !== context.lastColumn)
	)
		return 'tableTab.selectCell';
}

type SyncTableControls = typeof import('./table-ribbon-state').syncTableControls;
const panelSync = new WeakMap<HTMLElement, SyncTableControls>();

export function syncTableTab(core: EditorCore, measure = false): void {
	const { toolbar } = core.shell;
	const view = core.targetView();
	if (!toolbar || !view) return;
	const context = tableTabContext(view.state);
	const sync = panelSync.get(toolbar);
	if (sync) {
		sync(core, toolbar, view, context, measure);
		return;
	}
	if (!context || toolbar.dataset.tablePanels) return;
	toolbar.dataset.tablePanels = 'loading';
	void loadTablePanels(toolbar)
		.then((sync) => {
			toolbar.dataset.tablePanels = 'ready';
			panelSync.set(toolbar, sync);
			if (core.shell.toolbar === toolbar) syncTableTab(core);
		})
		.catch((error: unknown) => {
			delete toolbar.dataset.tablePanels;
			core.host.reportError(error);
		});
}

export function tableCells(node: Node, tablePos: number) {
	const cells: { row: number; column: number; node: Node; pos: number }[] = [];
	node.forEach((row, rowOffset, rowIndex) =>
		row.forEach((cell, offset, column) =>
			cells.push({ row: rowIndex, column, node: cell, pos: tablePos + rowOffset + offset + 2 }),
		),
	);
	return cells;
}
