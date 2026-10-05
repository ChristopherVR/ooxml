// UI-COMMANDS entry: registers every command, dialog and ribbon tab on an editor context.
import { insertNowCommands } from './insert-now.js';
import type { Command } from '../commands.js';
import type { EditorContext } from '../context.js';
import { registerDialogs } from '../dialogs/index.js';
import { registerRibbonTabs } from '../ribbon/parts.js';
import { commandTabs } from '../ribbon/tabs/index.js';
import { alignmentCommands } from './alignment.js';
import { cellCommands } from './cells.js';
import { clipboardCommands, installFormatPainter } from './clipboard.js';
import { contextualCommands } from './contextual.js';
import { dataCommands } from './data.js';
import { editingCommands } from './editing.js';
import { fontCommands } from './font.js';
import { formulaCommands } from './formulas.js';
import { registerSheetIcons } from './icons.js';
import { insertCommands } from './insert.js';
import { numberCommands } from './number.js';
import { pageLayoutCommands } from './page-layout.js';
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
