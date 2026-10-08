/**
 * The controls at the right end of Excel's ribbon tab row: the Editing / Viewing selector, the
 * Comments toggle and Share, in that order, as Excel 365 places them. They sit in the shared
 * ribbon's `actions` slot; the title bar keeps the file name, search and the people in a shared
 * session.
 */
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { el } from './ribbon/controls';
import { ribbonIcon } from './ribbon/icons';
import type { XlsxCollaborationState } from './collaboration-types';

export interface RibbonActionsHandlers {
	setReadOnly(readOnly: boolean): void;
	isHidden(id: string): boolean;
	/** Sharing state, for Share's pressed state; absent leaves it unpressed. */
	collaboration?(): XlsxCollaborationState;
}

export interface RibbonActions {
	readonly element: HTMLElement;
	/** Re-reads read-only, sharing and the comments pane state. */
	refresh(): void;
	relocalize(): void;
}

export function createRibbonActions(
	ctx: EditorContext,
	handlers: RibbonActionsHandlers,
): RibbonActions {
	const doc = ctx.host.ownerDocument;
	const element = el(doc, 'div', 'xve-ribbon-actions');
	element.slot = 'actions';

	const modeBox = el(doc, 'label', 'xve-mode');
	const mode = el(doc, 'select', 'xve-mode-select');
	mode.append(new Option('', 'editing'), new Option('', 'viewing'));
	mode.addEventListener('change', () => handlers.setReadOnly(mode.value === 'viewing'));
	modeBox.append(ribbonIcon(doc, 'pencil', 16), mode);

	const labelled = (className: string, icon: string) => {
		const button = el(doc, 'button', className);
		button.type = 'button';
		const text = el(doc, 'span', 'xve-action-label');
		button.append(ribbonIcon(doc, icon, 16), text);
		return { button, text };
	};
	const comments = labelled('xve-icon-button xve-comments-button', 'comments');
	comments.button.addEventListener('click', () => void ctx.commands.run('review.show-comments'));
	const share = labelled('xve-share-button', 'share');
	share.button.addEventListener('click', () => void ctx.commands.run('file.share'));
	element.append(modeBox, comments.button, share.button);

	const relocalize = () => {
		mode.setAttribute('aria-label', ctx.t('Editing mode'));
		mode.title = ctx.t('Editing mode');
		mode.options[0]!.textContent = ctx.t('Editing');
		mode.options[1]!.textContent = ctx.t('Viewing');
		comments.text.textContent = ctx.t('Comments');
		comments.button.title = ctx.t('Show comments');
		share.text.textContent = ctx.t('Share');
		share.button.title = ctx.t('Share');
		refresh();
	};
	const refresh = () => {
		mode.value = ctx.readOnly() ? 'viewing' : 'editing';
		share.button.hidden = !ctx.commands.get('file.share') || handlers.isHidden('file.share');
		share.button.setAttribute('aria-pressed', String(Boolean(handlers.collaboration?.().active)));
		const showComments = ctx.commands.get('review.show-comments');
		comments.button.hidden = !showComments;
		if (!showComments) return;
		let pressed = false;
		try {
			pressed = showComments.checked?.(ctx) ?? false;
		} catch {
			pressed = false;
		}
		comments.button.setAttribute('aria-pressed', String(pressed));
	};
	relocalize();
	return { element, refresh, relocalize };
}
