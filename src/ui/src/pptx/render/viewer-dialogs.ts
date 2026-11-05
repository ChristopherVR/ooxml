import {
	alertDialog,
	confirmDialog,
	promptDialog,
	type MessageDialogHost,
} from '../../dialog/message-dialogs';
import type { Translator } from '../i18n/translator';
import { shouldConfirmExternalHyperlink } from './options/viewer-options-apply';
import type { ViewerOptions } from './options/viewer-options';
import { safeOpenUrl } from './hyperlink-security';

/**
 * The in-app dialogs every PowerPoint binding uses instead of the browser's `alert`, `confirm`
 * and `prompt`. Each takes the binding's translator (`t`) and the element the dialog should theme
 * from (its viewer root); all strings are the shared `pptx.*` keys, so the five bindings ask the
 * same questions in the same words.
 */
export interface ViewerDialogContext {
	t: Translator;
	/** Viewer root (or any node inside it): the dialog mounts where the host's tokens apply. */
	host?: MessageDialogHost;
}

/** The Trust Center question before an external link opens. Resolves true to follow it. */
export function confirmExternalHyperlinkDialog(
	url: string,
	{ t, host }: ViewerDialogContext,
): Promise<boolean> {
	return confirmDialog({
		title: t('pptx.dialog.hyperlinkTitle'),
		message: `${t('pptx.dialog.hyperlinkMessage')}\n\n${url}`,
		confirmLabel: t('pptx.dialog.hyperlinkOpen'),
		cancelLabel: t('pptx.common.cancel'),
		host,
	});
}

/**
 * Open an external link, first asking when Options > Trust Center says to. The question is
 * asynchronous, so the click that started it must already have been cancelled (a link's default
 * navigation is not enough); the link opens only after the user agrees.
 */
export async function openExternalHyperlinkConfirmed(
	options: ViewerOptions,
	url: string,
	context: ViewerDialogContext,
	target?: string,
): Promise<boolean> {
	if (
		shouldConfirmExternalHyperlink(options, url) &&
		!(await confirmExternalHyperlinkDialog(url, context))
	)
		return false;
	safeOpenUrl(url, target);
	return true;
}

/** A destructive confirmation (delete, clear): the confirm button reads "Delete" and is red. */
export function confirmDeleteDialog(
	message: string,
	{ t, host }: ViewerDialogContext,
): Promise<boolean> {
	return confirmDialog({
		title: t('pptx.dialog.confirmTitle'),
		message,
		confirmLabel: t('pptx.common.delete'),
		cancelLabel: t('pptx.common.cancel'),
		tone: 'danger',
		host,
	});
}

/** A plain confirmation whose confirm button reads "OK". */
export function confirmActionDialog(
	message: string,
	{ t, host }: ViewerDialogContext,
): Promise<boolean> {
	return confirmDialog({
		title: t('pptx.dialog.confirmTitle'),
		message,
		confirmLabel: t('pptx.common.ok'),
		cancelLabel: t('pptx.common.cancel'),
		host,
	});
}

/** Ask for a name. Resolves the entered text, or null when the user cancels. */
export function promptNameDialog(
	message: string,
	defaultValue: string,
	{ t, host }: ViewerDialogContext,
): Promise<string | null> {
	return promptDialog({
		title: t('pptx.dialog.inputTitle'),
		message,
		defaultValue,
		confirmLabel: t('pptx.common.ok'),
		cancelLabel: t('pptx.common.cancel'),
		host,
	});
}

/** Report a failed action; resolves when the user dismisses it. */
export function alertNoticeDialog(
	message: string,
	{ t, host }: ViewerDialogContext,
): Promise<void> {
	return alertDialog({
		title: t('pptx.dialog.noticeTitle'),
		message,
		confirmLabel: t('pptx.common.ok'),
		host,
	});
}
