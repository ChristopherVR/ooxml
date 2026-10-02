// Registers the rule dialogs: conditional formatting (quick rules, rule editor, manager), data
// validation, sort and hyperlink.
import type { EditorContext } from '../context.js';
import { cfManagerDialog } from './cf-manager.js';
import { cfQuickDialog } from './cf-quick.js';
import { cfRuleDialog } from './cf-rule.js';
import { dataValidationDialog } from './data-validation.js';
import { hyperlinkDialog } from './hyperlink.js';
import { sortDialog } from './sort-dialog.js';

export function registerRuleDialogs(ctx: EditorContext): void {
	const d = ctx.dialogs;
	d.register('cf-quick', (c, props) => cfQuickDialog(c, props));
	d.register('cf-rule', (c, props) => cfRuleDialog(c, props));
	d.register('cf-manager', (c) => cfManagerDialog(c));
	d.register('data-validation', (c) => dataValidationDialog(c));
	d.register('sort', (c) => sortDialog(c));
	d.register('hyperlink', (c) => hyperlinkDialog(c));
}
