// Registers the navigation, names and information dialogs.
import type { EditorContext } from '../context.js';
import { openCommentsList } from './comments-list.js';
import { openCreateNames } from './create-names.js';
import { type DefineNameProps, openDefineName } from './define-name.js';
import { type FindReplaceProps, openFindReplace } from './find-replace.js';
import { openGoTo, openGoToSpecial } from './go-to.js';
import { openFeatureStatus, openShortcutHelp } from './help-dialogs.js';
import { type InsertFunctionProps, openInsertFunction } from './insert-function.js';
import { openNameManager } from './name-manager.js';
import { openStatistics } from './statistics.js';

const obj = <T>(props: unknown): T =>
	props && typeof props === 'object' ? (props as T) : ({} as T);

export function registerNavigationDialogs(ctx: EditorContext): void {
	const d = ctx.dialogs;
	d.register('find-replace', (c, p) => openFindReplace(c, obj<FindReplaceProps>(p)));
	d.register('go-to', (c) => openGoTo(c));
	d.register('go-to-special', (c) => openGoToSpecial(c));
	d.register('insert-function', (c, p) => openInsertFunction(c, obj<InsertFunctionProps>(p)));
	d.register('name-manager', (c) => openNameManager(c));
	d.register('define-name', (c, p) => openDefineName(c, obj<DefineNameProps>(p)));
	d.register('create-names', (c) => openCreateNames(c));
	d.register('workbook-statistics', (c) => openStatistics(c));
	d.register('comments-list', (c) => openCommentsList(c));
	d.register('shortcut-help', (c) => openShortcutHelp(c));
	d.register('feature-status', (c) => openFeatureStatus(c));
}
