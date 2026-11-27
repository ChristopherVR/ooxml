import { setTableLook, type Table } from 'ooxml-core/docx';
import { dispatchTableEdit } from './table-edit-transaction';
import { DEFAULT_TABLE_LOOK, type WordTableTools } from 'ooxml-core/docx/ui';
import { twipsFromPixels } from 'ooxml-core/units';
import type { EditorCore } from './editor-core';
import { focusView } from './focus-view';
import { openColorGridPopover } from './ribbon-color-grid';
import { runToInlineNodes } from './run-adapter';
import { applyCellFill } from './table-shading';
import {
	DEFAULT_PEN,
	setCellBorders,
	tableBorderContext,
	type BorderPen,
} from './table-border-commands';
import { toggleTableDraw } from './table-draw';
import type { TableTabAction } from './table-ribbon';
import { openTableStyles } from './table-style-gallery';
import { freeformTableReason, tableTabContext, tableTabReason } from './table-tab-context';
import { openTableConvert, openTableFormula, openTableSort } from './table-tab-dialogs';
import {
	alignTableCells,
	currentTable,
	patchTableNode,
	reorderTableRows,
	replaceTableWithText,
	selectTableParts,
} from './table-tab-edits';

const pens = new WeakMap<EditorCore, BorderPen>();

function setPen(
	core: EditorCore,
	pen: BorderPen,
	action: Extract<TableTabAction, { key: 'pen' }>,
): void {
	const { property, value } = action;
	if (property === 'style') {
		const next =
			value === 'grid'
				? DEFAULT_PEN
				: value === 'accent'
					? { style: 'double', sizeEighthPoints: 12, color: '#4472C4' }
					: { ...pen, style: value };
		pens.set(core, next);
	} else if (property === 'sizeEighthPoints')
		pens.set(core, { ...pen, sizeEighthPoints: Number(value) });
	else {
		const { color: _old, ...base } = pen;
		pens.set(core, value === 'auto' ? base : { ...base, color: value });
	}
}

export function runTableTabAction(
	core: EditorCore,
	action: TableTabAction,
	tools: WordTableTools,
	opener?: HTMLElement,
): void {
	if (action.key === 'draw' && action.value !== 'paint') {
		// Turning a tool off is always allowed; turning it on needs an editable body.
		if (core.shell.canvas?.dataset.tableDraw === action.value || !freeformTableReason(core))
			toggleTableDraw(core, action.value, pens.get(core) ?? DEFAULT_PEN);
		return;
	}
	const view = core.targetView();
	const context = view && tableTabContext(view.state);
	if (!view || !context || tableTabReason(core, action.key, context)) return;
	const { top, bottom, firstColumn, lastColumn, node, pos } = context;
	const minimumColumn =
		top === bottom
			? Math.min(firstColumn, lastColumn)
			: firstColumn === lastColumn
				? firstColumn
				: 0;
	const maximumColumn =
		top === bottom
			? Math.max(firstColumn, lastColumn)
			: firstColumn === lastColumn
				? firstColumn
				: (node.firstChild?.childCount ?? 1) - 1;
	const pen = pens.get(core) ?? DEFAULT_PEN;
	if (action.key === 'gridlines') {
		core.shell.canvas?.toggleAttribute('data-table-gridlines');
		core.refreshControls();
		return;
	}
	if (action.key === 'select') {
		selectTableParts(view, context, action.value);
		focusView(view);
		return;
	}
	if (action.key === 'margins') return core.formatDialogs.open('tableProperties', opener);
	if (action.key === 'pen') return setPen(core, pen, action);
	if (action.key === 'draw') return toggleTableDraw(core, action.value, pen);
	if (action.key === 'borders') {
		setCellBorders(view, action.value, pen);
		return;
	}
	if (action.key === 'align') return alignTableCells(view, context, action);
	const table = currentTable(core.model, context);
	if (!table) return;
	const selection = view.state.selection;
	const valid = () =>
		core.targetView() === view &&
		view.state.selection.eq(selection) &&
		!tableTabReason(core, action.key, tableTabContext(view.state));
	const apply = (updated: Table) => {
		const tr = patchTableNode(
			view.state.tr,
			pos,
			node,
			updated,
			core.model.tableStyles,
			core.model.theme,
		);
		dispatchTableEdit(view, tr);
	};
	switch (action.key) {
		case 'styles':
			openTableStyles(core, view, context, table, tools);
			break;
		case 'shading': {
			const anchor = core.shell.toolbar?.querySelector<HTMLElement>(
				'[data-table-command="shading"]',
			);
			if (!anchor) break;
			openColorGridPopover(
				anchor,
				(value) => {
					const current = tableBorderContext(view.state);
					if (!current || !valid()) return;
					const tr = view.state.tr;
					applyCellFill(tr, current, 'cell', value === 'none' ? null : value);
					dispatchTableEdit(view, tr);
				},
				{ noneLabel: 'No color' },
			);
			break;
		}
		case 'look': {
			const look = table.look ?? DEFAULT_TABLE_LOOK;
			apply(setTableLook(table, { ...look, [action.flag]: !look[action.flag] }));
			break;
		}
		case 'autoFit':
			apply(tools.autoFitTable(table, action.value));
			break;
		case 'repeat':
			apply(tools.repeatTableRows(table, top, bottom));
			break;
		case 'size':
		case 'distributeColumns':
		case 'distributeRows': {
			const width = twipsFromPixels(core.contentWidth());
			const value = action.key === 'size' ? Math.round(action.value * 1440) : undefined;
			if (
				action.key === 'distributeColumns' ||
				(action.key === 'size' && action.dimension === 'width')
			) {
				apply(tools.resizeTableColumns(table, minimumColumn, maximumColumn, width, value));
			} else {
				const tableDOM = view.nodeDOM(pos);
				const rows = tableDOM instanceof HTMLTableElement ? tableDOM.rows : null;
				const measured = table.rows.map(
					(_, index) =>
						table.rowProperties?.[index]?.heightTwips ??
						Math.max(1, Math.round((rows?.[index]?.offsetHeight || 20) * 15)),
				);
				apply(tools.resizeTableRows(table, top, bottom, measured, value));
			}
			break;
		}
		case 'formula':
			openTableFormula(core, top > 0 || firstColumn === 0, (instruction) => {
				if (!valid()) return false;
				const active = currentTable(core.model, context);
				const runs = active && tools.tableFormulaRuns(active, top, firstColumn, instruction);
				if (!runs) return false;
				const nodes = runs.flatMap((run) => runToInlineNodes(run));
				const { from, to } = view.state.selection;
				dispatchTableEdit(view, view.state.tr.replaceWith(from, to, nodes).scrollIntoView());
				return true;
			});
			break;
		case 'convert':
			openTableConvert(core, (separator) => {
				if (!valid()) return false;
				replaceTableWithText(view, context, tools.tableToText(table, separator));
				return true;
			});
			break;
		case 'sort': {
			const width = node.firstChild?.childCount ?? 0;
			const count = table.rowProperties?.findIndex((row) => !row.header) ?? 0;
			openTableSort(core, width, firstColumn, count < 0 ? table.rows.length : count, (settings) => {
				if (!valid()) return false;
				reorderTableRows(view, context, table, tools.sortTableRows(table, settings));
				return true;
			});
			break;
		}
	}
}
