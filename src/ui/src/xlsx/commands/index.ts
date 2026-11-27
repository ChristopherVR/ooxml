// UI-COMMANDS entry: registers every command, dialog and ribbon tab on an editor context.
import { insertNowCommands } from './insert-now';
import type { Command } from 'ooxml-core/xlsx/ui';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { registerDialogs } from '../dialogs/index';
import { registerRibbonTabs } from 'ooxml-core/xlsx/ui';
import { commandTabs } from '../ribbon/tabs/index';
import { alignmentCommands } from 'ooxml-core/xlsx/ui';
import { cellCommands } from 'ooxml-core/xlsx/ui';
import { clipboardCommands, installFormatPainter } from 'ooxml-core/xlsx/ui';
import { contextualCommands } from 'ooxml-core/xlsx/ui';
import { dataCommands, whatIfCommands } from 'ooxml-core/xlsx/ui';
import { editingCommands } from 'ooxml-core/xlsx/ui';
import { fontCommands } from 'ooxml-core/xlsx/ui';
import { formulaCommands } from 'ooxml-core/xlsx/ui';
import { registerSheetIcons } from './register-icons';
import { arrangeCommands } from './arrange';
import { drawingFormatCommands } from './drawing-format';
import { insertShapeCommands } from './insert-shapes';
import { drawCommands } from './draw';
import { addInCommands } from './add-in-selection';
import { insertCommands } from './insert';
import { insertAnnotationCommands } from './insert-annotations';
import { sheetPictureCommands } from './sheet-picture';
import { localObjectCommands } from './local-object';
import { sparklineCommands } from './sparklines';
import { pivotCommands } from './pivots';
import { pivotAnalyzeCommands } from './pivot-context';
import { pivotDesignCommands } from './pivot-design';
import { slicerCommands } from './slicers';
import { numberCommands } from 'ooxml-core/xlsx/ui';
import { pageLayoutCommands } from 'ooxml-core/xlsx/ui';
import { pageBreakCommands } from 'ooxml-core/xlsx/ui';
import { reviewCommands } from './review';
import { styleCommands } from './styles';
import { viewCommands } from './view';
import { chartAreaCommand, chartSeriesCommand } from './chart-format';
import { chartFormatTabCommands } from './chart-format-tab';
import { smartArtCommands } from '../smartart-controller';

/** Every UI-COMMANDS command (fresh objects; safe to register on several editors). */
export function allCommands(): Command[] {
	return [
		...clipboardCommands(),
		...fontCommands(),
		...alignmentCommands(),
		...numberCommands(),
		...styleCommands(),
		...cellCommands(),
		...editingCommands(),
		...insertNowCommands(),
		...insertCommands(),
		...addInCommands(),
		...insertShapeCommands(),
		...drawCommands(),
		...arrangeCommands(),
		...drawingFormatCommands(),
		...insertAnnotationCommands(),
		...sheetPictureCommands(),
		...localObjectCommands(),
		...sparklineCommands(),
		...pivotCommands(),
		...pivotAnalyzeCommands(),
		...pivotDesignCommands(),
		...slicerCommands(),
		...pageLayoutCommands(),
		...pageBreakCommands(),
		...formulaCommands(),
		...dataCommands(),
		...whatIfCommands(),
		...reviewCommands(),
		...viewCommands(),
		...contextualCommands(),
		chartSeriesCommand(),
		chartAreaCommand(),
		...chartFormatTabCommands(),
		...smartArtCommands(),
	];
}

/**
 * Registers the commands, dialogs and ribbon tabs on `ctx`. Returns a disposer for the listeners
 * it adds (the format painter).
 */
export function installCommands(ctx: EditorContext): () => void {
	registerSheetIcons();
	ctx.commands.registerAll(allCommands());
	registerDialogs(ctx);
	registerRibbonTabs(commandTabs());
	return installFormatPainter(ctx);
}

export { commandTabs };
