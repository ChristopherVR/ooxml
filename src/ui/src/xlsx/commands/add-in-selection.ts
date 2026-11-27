import { getCell } from 'ooxml-core/xlsx';
import { target, sheetLocked, type Command, type EditorContext } from 'ooxml-core/xlsx/ui';
import { openOfficeAddIns, type AddInDialogMode } from '../../ribbon/add-in-dialog';
import { syncInstalledRibbonTabs } from '../../ribbon/add-in-tabs';
import type { OfficeAddInSelection } from '../../ribbon/add-in-runtime-settings';

/** Reads the workbook shown when each request arrives, so a pane opened at launch keeps working. */
export function excelAddInSelection(ctx: EditorContext): OfficeAddInSelection {
	const singleCell = () => {
		const selected = target(ctx);
		if (!selected || selected.ranges.length !== 1) return;
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

export function addInCommands(): Command[] {
	const launch = (id: string, label: string, mode: AddInDialogMode): Command => ({
		id,
		label,
		icon: 'addIns',
		run: (ctx) => {
			void openOfficeAddIns(
				ctx.root,
				mode,
				(text) => ctx.t(text),
				'excel',
				ctx.locale(),
				() => excelAddInSelection(ctx),
				syncInstalledRibbonTabs,
			);
		},
	});
	return [
		launch('home.add-ins', 'Add-ins', 'myAddIns'),
		launch('insert.get-add-ins', 'Get Add-ins', 'getAddIns'),
		launch('insert.my-add-ins', 'My Add-ins', 'myAddIns'),
	];
}
