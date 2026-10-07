// Registers the navigation, names and information dialogs.
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { openCommentsList } from './comments-list';
import { openCreateNames } from './create-names';
import { type DefineNameProps, openDefineName } from './define-name';
import { type FindReplaceProps, openFindReplace } from './find-replace';
import { openGoTo, openGoToSpecial } from './go-to';
import { openFeatureStatus, openShortcutHelp } from './help-dialogs';
import { type InsertFunctionProps, openInsertFunction } from './insert-function';
import { openNameManager } from './name-manager';
import { openStatistics } from './statistics';

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
