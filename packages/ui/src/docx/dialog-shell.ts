import { defineDialog } from '../controls';
import {
	localizeElement,
	translate,
	type EditorLocale,
	type LocalizationKey,
} from './localization';

type DialogElement = HTMLElement & { open: boolean; heading: string };

/**
 * The frame every format dialog shares: the shared modal `office-ui-dialog` (focus trap, Escape,
 * backdrop, heading, return of focus). The dialog's fields are its light-DOM children, styled by
 * the editor, and its buttons go in `dialogActions`. `title` is the English heading; it is
 * translated by `localizeDialog`.
 */
export function createDialogShell(title: string, className: string): HTMLElement {
	defineDialog();
	const element = document.createElement('office-ui-dialog') as DialogElement;
	element.className = `dve-dialog ${className}`.trim();
	element.dataset.title = title;
	element.heading = title;
	element.setAttribute('close-label', 'Close');
	return element;
}

/** The action row, placed in the dialog's footer slot. */
export function dialogActions(...buttons: HTMLElement[]): HTMLDivElement {
	const actions = document.createElement('div');
	actions.className = 'dve-dialog-actions';
	actions.slot = 'footer';
	actions.append(...buttons);
	return actions;
}

export const isDialogOpen = (element: HTMLElement): boolean => (element as DialogElement).open;

export function setDialogOpen(element: HTMLElement, open: boolean): void {
	(element as DialogElement).open = open;
}

/**
 * Runs `onDismiss` when the user closes the dialog with Escape, the close button or the backdrop,
 * so the dialog can clean up the way its own Cancel button does.
 */
export function onDialogDismiss(element: HTMLElement, onDismiss: () => void): void {
	element.addEventListener('office-dialog-close', (event) => {
		event.preventDefault();
		onDismiss();
	});
}

/** Translates the dialog's fields (light DOM) and its heading. */
export function localizeDialog(element: HTMLElement, locale: EditorLocale): void {
	localizeElement(element, locale);
	const title = element.dataset.title;
	if (title) (element as DialogElement).heading = translate(locale, title as LocalizationKey);
	element.setAttribute('close-label', translate(locale, 'Close'));
}
