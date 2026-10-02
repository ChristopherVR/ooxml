// Registers every UI-COMMANDS dialog on the editor's dialog registry (opened by name).
import type { EditorContext } from '../context.js';
import { registerFormatCellsDialogs } from './format-cells/index.js';
import { registerNavigationDialogs } from './register-navigation.js';
import { registerRuleDialogs } from './register-rules.js';
import { registerSimpleDialogs } from './simple.js';
import { registerToolDialogs } from './register-tools.js';

export function registerDialogs(ctx: EditorContext): void {
	registerSimpleDialogs(ctx);
	registerNavigationDialogs(ctx);
	registerToolDialogs(ctx);
	registerFormatCellsDialogs(ctx);
	registerRuleDialogs(ctx);
}
