/**
 * The controls at the right end of Word's ribbon tab row: the Editing / Viewing selector, the
 * Comments toggle and Share, in that order, as Word 365 places them (Excel uses the same
 * `office-ui-ribbon-actions` element). They sit in the shared ribbon's `actions` slot; the title
 * bar keeps the quick access toolbar, the file name and "Tell me".
 */
import { defineRibbonActions, type OfficeUiRibbonActions } from '../ribbon/ribbon-actions';
import { translate, type EditorLocale } from './localization';

export interface RibbonActionsHandlers {
	setReadOnly(readOnly: boolean): void;
	toggleComments(): void;
	share(): void;
}

export interface RibbonActions {
	readonly element: OfficeUiRibbonActions;
	setReadOnly(readOnly: boolean): void;
	setCommentsOpen(open: boolean): void;
	/** Share reads pressed while a collaboration session is live. */
	setSharing(active: boolean): void;
	relocalize(locale: EditorLocale): void;
}

export function createRibbonActions(handlers: RibbonActionsHandlers): RibbonActions {
	defineRibbonActions();
	const element = document.createElement('office-ui-ribbon-actions') as OfficeUiRibbonActions;
	element.className = 'dve-ribbon-actions';
	element.slot = 'actions';
	element.mode = 'editing';
	element.addEventListener('office-ribbon-mode', (event) =>
		handlers.setReadOnly((event as CustomEvent<{ mode: string }>).detail.mode === 'viewing'),
	);
	element.addEventListener('office-ribbon-comments', () => handlers.toggleComments());
	element.addEventListener('office-ribbon-share', () => handlers.share());

	const relocalize = (locale: EditorLocale) => {
		element.modeLabel = translate(locale, 'Editing mode');
		element.modes = [
			{ value: 'editing', label: translate(locale, 'Editing') },
			{ value: 'viewing', label: translate(locale, 'Viewing') },
		];
		element.commentsLabel = translate(locale, 'Comments');
		element.commentsTitle = translate(locale, 'Show comments');
		element.shareLabel = translate(locale, 'Share');
		element.shareTitle = translate(locale, 'Share');
	};
	relocalize('en');
	return {
		element,
		setReadOnly(readOnly) {
			element.mode = readOnly ? 'viewing' : 'editing';
		},
		setCommentsOpen(open) {
			element.commentsPressed = open;
		},
		setSharing(active) {
			element.sharePressed = active;
		},
		relocalize,
	};
}
