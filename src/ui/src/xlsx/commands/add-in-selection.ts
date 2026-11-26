import { getCell } from 'ooxml-core/xlsx';
import { target, sheetLocked, type EditorContext } from 'ooxml-core/xlsx/ui';
import type { OfficeAddInSelection } from '../../ribbon/add-in-runtime-settings';

export function excelAddInSelection(ctx: EditorContext): OfficeAddInSelection {
	const session = ctx.session();
	const singleCell = () => {
		const selected = target(ctx);
		if (!selected || selected.session !== session || selected.ranges.length !== 1) return;
		const { start, end } = selected.range;
		if (start.row === end.row && start.col === end.col) return selected;
	};
	return {
		readText() {
			const selected = singleCell();
			if (!selected) return;
			const { row, col } = selected.active;
			const value = getCell(selected.ws, row, col)?.value;
			if (value === null || value === undefined) return '';
			if (typeof value === 'string' || typeof value === 'number' || typeof value === 'boolean')
				return String(value);
		},
		writeText(text) {
			const selected = singleCell();
			if (!selected || ctx.readOnly() || sheetLocked(ctx)) return false;
			const { row, col } = selected.active;
			selected.session.setCellValue(selected.sheet, row, col, text);
			return true;
		},
	};
}
