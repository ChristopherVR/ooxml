// Registers every UI-COMMANDS dialog on the editor's dialog registry (opened by name).
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { registerFormatCellsDialogs } from './format-cells/index';
import { registerNavigationDialogs } from './register-navigation';
import { registerRuleDialogs } from './register-rules';
import { registerSimpleDialogs } from './simple';
import { registerToolDialogs } from './register-tools';

export function registerDialogs(ctx: EditorContext): void {
	registerSimpleDialogs(ctx);
	registerNavigationDialogs(ctx);
	registerToolDialogs(ctx);
	registerFormatCellsDialogs(ctx);
	registerRuleDialogs(ctx);
}
