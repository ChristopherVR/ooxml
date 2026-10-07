// Registers the rule dialogs: conditional formatting (quick rules, rule editor, manager), data
// validation, sort and hyperlink.
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { cfManagerDialog } from './cf-manager';
import { cfQuickDialog } from './cf-quick';
import { cfRuleDialog } from './cf-rule';
import { dataValidationDialog } from './data-validation';
import { hyperlinkDialog } from './hyperlink';
import { sortDialog } from './sort-dialog';

export function registerRuleDialogs(ctx: EditorContext): void {
	const d = ctx.dialogs;
	d.register('cf-quick', (c, props) => cfQuickDialog(c, props));
	d.register('cf-rule', (c, props) => cfRuleDialog(c, props));
	d.register('cf-manager', (c) => cfManagerDialog(c));
	d.register('data-validation', (c) => dataValidationDialog(c));
	d.register('sort', (c) => sortDialog(c));
	d.register('hyperlink', (c) => hyperlinkDialog(c));
}
