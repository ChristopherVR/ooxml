// UI-COMMANDS entry: registers every command, dialog and ribbon tab on an editor context.
import { insertNowCommands } from './insert-now.js';
import type { Command } from 'ooxml-core/xlsx/ui';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { registerDialogs } from '../dialogs/index.js';
import { registerRibbonTabs } from 'ooxml-core/xlsx/ui';
import { commandTabs } from '../ribbon/tabs/index.js';
import { alignmentCommands } from 'ooxml-core/xlsx/ui';
import { cellCommands } from 'ooxml-core/xlsx/ui';
import { clipboardCommands, installFormatPainter } from 'ooxml-core/xlsx/ui';
import { contextualCommands } from 'ooxml-core/xlsx/ui';
import { dataCommands } from 'ooxml-core/xlsx/ui';
import { editingCommands } from 'ooxml-core/xlsx/ui';
import { fontCommands } from 'ooxml-core/xlsx/ui';
import { formulaCommands } from 'ooxml-core/xlsx/ui';
import { registerSheetIcons } from './register-icons.js';
import { insertCommands } from './insert.js';
import { numberCommands } from 'ooxml-core/xlsx/ui';
import { pageLayoutCommands } from 'ooxml-core/xlsx/ui';
import { reviewCommands } from './review.js';
import { styleCommands } from './styles.js';
import { viewCommands } from './view.js';

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
		...pageLayoutCommands(),
		...formulaCommands(),
		...dataCommands(),
		...reviewCommands(),
		...viewCommands(),
		...contextualCommands(),
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
