import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { watchOfficeAddIns } from '../ribbon/add-in-dialog';
import { excelAddInSelection } from './commands/add-in-selection';

/** Restores the workbook editor's mounted add-ins and keeps them in step with its preferences. */
export function watchExcelAddIns(ctx: EditorContext): () => void {
	return watchOfficeAddIns(
		ctx.root,
		(label) => ctx.t(label),
		'excel',
		() => ctx.locale(),
		() => excelAddInSelection(ctx),
	);
}
