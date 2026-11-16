import { TABLE_CELL_SECTIONS, withCellStyle } from 'ooxml-ui/pptx';
import type { PptxTableCellStyle, PptxTableData } from 'pptx-viewer-core';
import React from 'react';

import { useTranslation } from '../../i18n/use-translation';
import type { TableCellEditorState } from '../../types';
import { cn } from '../../utils';
import { HEADING, CARD, BTN } from './inspector-pane-constants';
import { InspectorFields } from './InspectorFields';
import {
	computeMergeCellRight,
	computeMergeCellDown,
	computeSplitCell,
} from './table-cell-merge-helpers';
import { TableCellAdvancedFill } from './TableCellAdvancedFill';
import { TableCellColorField } from './TableCellColorField';

// ---------------------------------------------------------------------------
// Props
// ---------------------------------------------------------------------------

interface TableCellFormattingPanelProps {
	tableData: PptxTableData;
	tableEditorState: TableCellEditorState;
	canEdit: boolean;
	onUpdateTableData: (patch: Partial<PptxTableData>) => void;
	onUpdateMergeRows: (rows: PptxTableData['rows']) => void;
}

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

export function TableCellFormattingPanel({
	tableData: td,
	tableEditorState,
	canEdit,
	onUpdateTableData,
	onUpdateMergeRows,
}: TableCellFormattingPanelProps): React.ReactElement | null {
	const { t } = useTranslation();
	const { rowIndex, columnIndex } = tableEditorState;
	const cell = td.rows[rowIndex]?.cells[columnIndex];
	if (!cell) {
		return null;
	}
	const cs: PptxTableCellStyle = cell.style ?? {};

	const updateCellStyle = (updates: Partial<PptxTableCellStyle>) => {
		const newRows = td.rows.map((row, ri) => {
			if (ri !== rowIndex) {
				return row;
			}
			return {
				...row,
				cells: row.cells.map((c, ci) => (ci === columnIndex ? withCellStyle(c, updates) : c)),
			};
		});
		onUpdateTableData({ rows: newRows });
	};

	return (
		<div className={CARD}>
			<div className={HEADING}>
				{t('pptx.table.cell', { row: rowIndex + 1, col: columnIndex + 1 })}
			</div>
			<div className='space-y-2'>
				{/* Colors */}
				<div className='grid grid-cols-2 gap-1.5'>
					<TableCellColorField
						label={t('pptx.table.color')}
						value={cs.color}
						fallback='#000000'
						selectedRef={cs.colorRef}
						disabled={!canEdit}
						onCommit={(hex, ref) => updateCellStyle({ color: hex, colorRef: ref })}
					/>
					<TableCellColorField
						label={t('pptx.table.background')}
						value={cs.backgroundColor}
						fallback='#ffffff'
						selectedRef={cs.backgroundColorRef}
						disabled={!canEdit}
						onCommit={(hex, ref) =>
							updateCellStyle({ backgroundColor: hex, backgroundColorRef: ref })
						}
					/>
				</div>

				{/* Advanced Cell Fill */}
				<TableCellAdvancedFill
					cellStyle={cs}
					canEdit={canEdit}
					onUpdateCellStyle={updateCellStyle}
				/>

				{/* Size, emphasis, alignment, margins and borders: shared descriptors. */}
				<div className='grid grid-cols-2 gap-1.5 text-[11px]'>
					<InspectorFields
						sections={TABLE_CELL_SECTIONS}
						target={cs}
						disabled={!canEdit}
						onPatch={updateCellStyle}
					/>
				</div>

				{/* Merge / Split */}
				<div className='grid grid-cols-3 gap-1'>
					<button
						type='button'
						disabled={!canEdit}
						className={cn(BTN, 'text-center')}
						onClick={() => {
							const newRows = computeMergeCellRight(td, rowIndex, columnIndex);
							if (newRows) {
								onUpdateMergeRows(newRows);
							}
						}}
					>
						{t('pptx.table.mergeRight')}
					</button>
					<button
						type='button'
						disabled={!canEdit}
						className={cn(BTN, 'text-center')}
						onClick={() => {
							const newRows = computeMergeCellDown(td, rowIndex, columnIndex);
							if (newRows) {
								onUpdateMergeRows(newRows);
							}
						}}
					>
						{t('pptx.table.mergeDown')}
					</button>
					<button
						type='button'
						disabled={!canEdit}
						className={cn(BTN, 'text-center')}
						onClick={() => {
							const newRows = computeSplitCell(td, rowIndex, columnIndex);
							if (newRows) {
								onUpdateMergeRows(newRows);
							}
						}}
					>
						{t('pptx.table.split')}
					</button>
				</div>
			</div>
		</div>
	);
}
