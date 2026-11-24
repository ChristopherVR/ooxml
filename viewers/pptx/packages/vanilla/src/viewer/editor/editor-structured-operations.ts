import { getCellText, resolveTableCellEditEnd, withCellText } from 'ooxml-ui/pptx';
import type { TableCellEditEnd, TableCellEditOutcome } from 'ooxml-ui/pptx';
import type { PptxElement, TablePptxElement } from 'pptx-viewer-core';
import { hasTextProperties } from 'pptx-viewer-core';

import type { Store, ViewerState } from '../state';
import {
	findActiveElement,
	mapActiveElement,
	replaceActiveElements,
} from './editor-active-elements';

export interface StructuredEditorOperations {
	commitTableCell(id: string, row: number, column: number, text: string): void;
	endTableCellEdit(
		id: string,
		row: number,
		column: number,
		end: TableCellEditEnd,
	): TableCellEditOutcome | undefined;
	/** Replace a table element as one undoable step (a paste into selected cells). */
	replaceTableElement(next: TablePptxElement): void;
	updateEquation(id: string, omml: Record<string, unknown>): void;
}

export function createStructuredEditorOperations(deps: {
	store: Store<ViewerState>;
	pushHistory(): void;
	commitChange(): void;
	/** Options > Proofing > AutoCorrect, applied to committed table-cell text. */
	transformCommittedText?: (text: string, previous?: string) => string;
}): StructuredEditorOperations {
	return {
		commitTableCell(id, row, column, rawText) {
			const state = deps.store.get();
			const target = findActiveElement(state, id);
			if (target?.type !== 'table' || !target.tableData?.rows[row]?.cells[column]) {
				return;
			}
			const text = deps.transformCommittedText
				? deps.transformCommittedText(rawText, getCellText(target, row, column))
				: rawText;
			deps.pushHistory();
			const rows = target.tableData.rows.map((tableRow, rowIndex) => ({
				...tableRow,
				cells: tableRow.cells.map((cell, cellIndex) =>
					// Shared `withCellText`, not a local `textRuns: undefined`. Vanilla
					// was the ONLY binding clearing the stale run model here, so it was
					// the only one whose cell edits painted; that made a local defence
					// out of a decision all five bindings need.
					rowIndex === row && cellIndex === column ? withCellText(cell, text) : cell,
				),
			}));
			deps.store.set(
				replaceActiveElements(
					state,
					mapActiveElement(state, id, (element) =>
						element.type === 'table'
							? { ...element, tableData: { ...element.tableData!, rows } }
							: element,
					),
				),
			);
			deps.commitChange();
		},
		replaceTableElement(next) {
			const state = deps.store.get();
			if (findActiveElement(state, next.id)?.type !== 'table') {
				return;
			}
			deps.pushHistory();
			deps.store.set(
				replaceActiveElements(
					state,
					mapActiveElement(state, next.id, () => next),
				),
			);
			deps.commitChange();
		},
		endTableCellEdit(id, row, column, end) {
			const state = deps.store.get();
			const target = findActiveElement(state, id);
			if (target?.type !== 'table') {
				return undefined;
			}
			// Shared: commits the runs, appends a row for Tab in the last cell,
			// fills cells from a pasted grid, and names the next cell.
			const outcome = resolveTableCellEditEnd(target, row, column, end, {
				...(deps.transformCommittedText ? { transformText: deps.transformCommittedText } : {}),
			});
			const next = outcome.element;
			const focusCell = outcome.edit ?? outcome.select;
			const cellPatch = focusCell
				? {
						selectedTableCell: { row: focusCell.row, column: focusCell.col },
						selectedTableCells: [{ row: focusCell.row, column: focusCell.col }],
					}
				: {};
			if (next) {
				deps.pushHistory();
				deps.store.set({
					...replaceActiveElements(
						state,
						mapActiveElement(state, id, () => next),
					),
					...cellPatch,
				});
				deps.commitChange();
			} else if (focusCell) {
				deps.store.set(cellPatch);
			}
			return outcome;
		},
		updateEquation(id, omml) {
			const state = deps.store.get();
			const target = findActiveElement(state, id);
			if (
				!target ||
				!hasTextProperties(target) ||
				!target.textSegments?.some((segment) => segment.equationXml)
			) {
				return;
			}
			deps.pushHistory();
			deps.store.set(
				replaceActiveElements(
					state,
					mapActiveElement(state, id, (element) =>
						hasTextProperties(element)
							? ({
									...element,
									textSegments: element.textSegments?.map((segment) => {
										if (!segment.equationXml) {
											return segment;
										}
										// Drop any captured `equationSourceXml`: it is the
										// OLD equation's untouched wrapper, and the writer
										// prefers it over `equationXml`, so leaving it in
										// place would keep re-emitting the equation the user
										// just edited away from.
										const { equationSourceXml: _dropped, ...rest } = segment;
										return { ...rest, equationXml: omml };
									}),
								} as PptxElement)
							: element,
					),
				),
			);
			deps.commitChange();
		},
	};
}
