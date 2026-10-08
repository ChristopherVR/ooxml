/**
 * The controls at the right end of Excel's ribbon tab row: the Editing / Viewing selector, the
 * Comments toggle and Share, in that order, as Excel 365 places them. The shared
 * `office-ui-ribbon-actions` element draws them (Word uses the same one) in the shared ribbon's
 * `actions` slot; the title bar keeps the file name, search and the people in a shared session.
 */
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import { defineRibbonActions, type OfficeUiRibbonActions } from '../ribbon/ribbon-actions';
import type { XlsxCollaborationState } from './collaboration-types';

export interface RibbonActionsHandlers {
	setReadOnly(readOnly: boolean): void;
	isHidden(id: string): boolean;
	/** Sharing state, for Share's pressed state; absent leaves it unpressed. */
	collaboration?(): XlsxCollaborationState;
}

export interface RibbonActions {
	readonly element: OfficeUiRibbonActions;
	/** Re-reads read-only, sharing and the comments pane state. */
	refresh(): void;
	relocalize(): void;
}

export function createRibbonActions(
	ctx: EditorContext,
	handlers: RibbonActionsHandlers,
): RibbonActions {
	const doc = ctx.host.ownerDocument;
	defineRibbonActions(doc.defaultView?.customElements);
	const element = doc.createElement('office-ui-ribbon-actions') as OfficeUiRibbonActions;
	element.className = 'xve-ribbon-actions';
	element.slot = 'actions';
	element.addEventListener('office-ribbon-mode', (event) =>
		handlers.setReadOnly((event as CustomEvent<{ mode: string }>).detail.mode === 'viewing'),
	);
	element.addEventListener('office-ribbon-comments', () => {
		void ctx.commands.run('review.show-comments');
	});
	element.addEventListener('office-ribbon-share', () => void ctx.commands.run('file.share'));

	const relocalize = () => {
		element.modeLabel = ctx.t('Editing mode');
		element.modes = [
			{ value: 'editing', label: ctx.t('Editing') },
			{ value: 'viewing', label: ctx.t('Viewing') },
		];
		element.commentsLabel = ctx.t('Comments');
		element.commentsTitle = ctx.t('Show comments');
		element.shareLabel = ctx.t('Share');
		element.shareTitle = ctx.t('Share');
		refresh();
	};
	const refresh = () => {
		element.mode = ctx.readOnly() ? 'viewing' : 'editing';
		element.noShare = !ctx.commands.get('file.share') || handlers.isHidden('file.share');
		element.sharePressed = Boolean(handlers.collaboration?.().active);
		const showComments = ctx.commands.get('review.show-comments');
		element.noComments = !showComments;
		if (!showComments) return;
		let pressed = false;
		try {
			pressed = showComments.checked?.(ctx) ?? false;
		} catch {
			pressed = false;
		}
		element.commentsPressed = pressed;
	};
	relocalize();
	return { element, refresh, relocalize };
}
