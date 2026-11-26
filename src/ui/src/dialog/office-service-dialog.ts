import { defineDialog, type OfficeUiDialog } from './dialog';

export type OfficeServiceText = (label: string) => string;

export function showOfficeServiceDialog(
	host: HTMLElement | ShadowRoot,
	heading: string,
	closeLabel: string,
) {
	const doc = host.ownerDocument;
	defineDialog(doc.defaultView?.customElements);
	const dialog = doc.createElement('office-ui-dialog') as OfficeUiDialog;
	dialog.heading = heading;
	dialog.closeLabel = closeLabel;
	dialog.style.setProperty('z-index', '1201');
	dialog.addEventListener('office-dialog-closed', (event) => {
		if (event.target === dialog) dialog.remove();
	});
	host.append(dialog);
	return dialog;
}

export function officeServiceLink(doc: Document, label: string, url: string): HTMLAnchorElement {
	const link = doc.createElement('a');
	link.textContent = label;
	link.href = url;
	link.target = '_blank';
	link.rel = 'noopener noreferrer';
	return link;
}
