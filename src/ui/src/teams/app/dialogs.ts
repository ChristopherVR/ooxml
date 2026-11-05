import { confirmDialog, promptDialog, type MessageDialogHost } from '../../dialog/message-dialogs';
import type { TeamsTranslate } from '../i18n/translator';

/** In-app replacements for the browser's confirm and prompt, worded by the Teams dictionary. */
export interface TeamsDialogContext {
	t: TeamsTranslate;
	/** Where the dialog mounts so it inherits the app's tokens (the app or an element inside it). */
	host?: MessageDialogHost;
}

/**
 * Ask a yes/no question. `confirmKey` names the confirm button (OK by default); `danger` styles it
 * as destructive and focuses Cancel first.
 */
export function confirmTeams(
	{ t, host }: TeamsDialogContext,
	message: string,
	options: { confirmKey?: 'common.ok' | 'common.discard' | 'chat.delete'; danger?: boolean } = {},
): Promise<boolean> {
	return confirmDialog({
		title: t('dialog.confirmTitle'),
		message,
		confirmLabel: t(options.confirmKey ?? 'common.ok'),
		cancelLabel: t('common.cancel'),
		tone: options.danger ? 'danger' : 'default',
		host,
	});
}

/** Ask for a name; resolves the trimmed text, or null when cancelled or left empty. */
export async function promptTeamsName(
	{ t, host }: TeamsDialogContext,
	message: string,
	defaultValue: string,
): Promise<string | null> {
	const entered = await promptDialog({
		title: t('dialog.inputTitle'),
		message,
		defaultValue,
		confirmLabel: t('common.ok'),
		cancelLabel: t('common.cancel'),
		host,
	});
	return entered?.trim() || null;
}
