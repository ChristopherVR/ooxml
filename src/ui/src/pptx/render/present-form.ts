import { officeServiceLink, showOfficeServiceDialog } from '../../dialog/office-service-dialog';
import type { Translator } from '../i18n/translator';

/** The caller validates the response URL. Opening this dialog does not contact Microsoft. */
export function presentMicrosoftForm(root: HTMLElement, url: string, t: Translator) {
	const doc = root.ownerDocument;
	const fullscreen = doc.fullscreenElement;
	const host = fullscreen instanceof HTMLElement && root.contains(fullscreen) ? fullscreen : root;
	const title = t('pptx.forms.Forms');
	const dialog = showOfficeServiceDialog(host, title, t('pptx.forms.Close'));
	dialog.addEventListener('click', (event) => event.stopPropagation());
	dialog.addEventListener('keydown', (event) => {
		if (event.key !== 'Escape') return;
		event.preventDefault();
		event.stopPropagation();
		dialog.close();
	});
	dialog.style.setProperty('--office-dialog-width', 'min(48rem, 95vw)');
	const actions = doc.createElement('div');
	const load = doc.createElement('button');
	load.type = 'button';
	load.textContent = t('pptx.forms.load');
	const external = officeServiceLink(doc, t('pptx.forms.openInBrowser'), url);
	actions.append(load, doc.createTextNode(' '), external);
	dialog.append(actions);
	load.addEventListener(
		'click',
		() => {
			const address = new URL(url);
			address.searchParams.set('embed', 'true');
			const frame = doc.createElement('iframe');
			frame.title = title;
			frame.tabIndex = 0;
			frame.referrerPolicy = 'no-referrer';
			frame.setAttribute('sandbox', 'allow-forms allow-scripts allow-same-origin');
			frame.style.cssText = 'display:block;width:100%;height:min(32rem,60vh);border:0';
			frame.src = address.href;
			dialog.prepend(frame);
			external.focus();
			load.remove();
		},
		{ once: true },
	);
	dialog.show();
	return dialog;
}
