import { defineDialog } from '../../dialog/dialog';
import type { OfficeUiDialog } from '../../dialog/dialog';
import { defineDialogFooter } from '../../dialog/dialog-footer';
import type { OfficeDialogFooterState } from '../../dialog/dialog-footer';

/** Translates a dictionary key; an untranslated key comes back unchanged. */
export type InsertDialogTranslate = (key: string) => string;

/** `translate(key)`, or the English fallback while the key has no entry. */
export function insertDialogText(
	translate: InsertDialogTranslate | undefined,
	key: string,
	fallback: string,
): string {
	const value = translate?.(key);
	return value && value !== key ? value : fallback;
}

type Dialog = OfficeUiDialog;
type Footer = HTMLElement & { state: OfficeDialogFooterState; focusAction(id: string): void };

/**
 * The frame of a PowerPoint insert dialog: the shared `office-ui-dialog` (heading, Escape,
 * backdrop, focus trap) around a shared picker, with Cancel and Insert in an
 * `office-ui-dialog-footer`. Rendered in the host element's light DOM so the dialog's focus
 * trap sees the picker's controls. `onClose` runs whenever the dialog closes itself.
 */
export interface InsertDialogFrame {
	dialog: Dialog;
	footer: Footer;
	sync(
		heading: string,
		closeLabel: string,
		cancel: string,
		insert: string,
		canInsert: boolean,
	): void;
}

export function createInsertDialogFrame(
	host: HTMLElement,
	body: HTMLElement,
	onInsert: () => void,
	onClose: () => void,
): InsertDialogFrame {
	const doc = host.ownerDocument;
	const registry = doc.defaultView?.customElements;
	defineDialog(registry);
	defineDialogFooter(registry);
	const dialog = doc.createElement('office-ui-dialog') as Dialog;
	const footer = doc.createElement('office-ui-dialog-footer') as Footer;
	footer.slot = 'footer';
	// Above the editor's popovers (1200), as the bindings' own dialogs were; wide enough for a gallery.
	dialog.style.setProperty('z-index', '1201');
	dialog.style.setProperty('--office-dialog-width', 'min(780px, calc(100vw - 32px))');
	dialog.style.setProperty('--office-dialog-max-width', 'calc(100vw - 32px)');
	dialog.append(body, footer);
	dialog.addEventListener('office-dialog-close', (event) => {
		if (!event.defaultPrevented) queueMicrotask(onClose);
	});
	footer.addEventListener('office-dialog-footer-request', (event) => {
		const id = (event as CustomEvent<{ id: string }>).detail.id;
		if (id === 'insert') onInsert();
		else dialog.close();
	});
	host.append(dialog);
	return {
		dialog,
		footer,
		sync(heading, closeLabel, cancel, insert, canInsert) {
			dialog.heading = heading;
			dialog.closeLabel = closeLabel;
			footer.state = {
				actions: [
					{ id: 'cancel', label: cancel },
					{ id: 'insert', label: insert, variant: 'primary', disabled: !canInsert },
				],
			};
		},
	};
}
