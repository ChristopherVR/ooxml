import {
	officeServiceLink,
	showOfficeServiceDialog,
	type OfficeServiceText,
} from '../dialog/office-service-dialog';

const RESPOND_PATH = '/pages/responsepage.aspx';

export function microsoftFormsResponseUrl(value: string): string | undefined {
	try {
		const url = new URL(value.trim());
		if (url.protocol !== 'https:' || url.username || url.password || url.port) return;
		if (
			(url.hostname === 'forms.office.com' || url.hostname === 'forms.cloud.microsoft') &&
			((url.pathname.toLowerCase() === RESPOND_PATH && !!url.searchParams.get('id')?.trim()) ||
				/^\/r\/[a-z\d]+\/?$/iu.test(url.pathname))
		) {
			return url.href;
		}
	} catch {
		return undefined;
	}
}

/** Collect a response link without fetching it or placing third-party content in the editor. */
export function pickMicrosoftForm(
	host: HTMLElement | ShadowRoot,
	t: OfficeServiceText,
): Promise<{ title: string; url: string } | undefined> {
	const doc = host.ownerDocument;
	const dialog = showOfficeServiceDialog(host, t('Forms'), t('Close'));
	const form = doc.createElement('form');
	const title = doc.createElement('input');
	title.value = t('Open form');
	title.maxLength = 120;
	const address = doc.createElement('input');
	address.type = 'url';
	address.required = true;
	const button = doc.createElement('button');
	button.type = 'submit';
	button.textContent = t('Insert link');
	const field = (label: string, input: HTMLInputElement) => {
		const row = doc.createElement('p');
		const wrapper = doc.createElement('label');
		wrapper.append(doc.createTextNode(t(label) + ' '), input);
		row.append(wrapper);
		return row;
	};
	form.append(
		officeServiceLink(doc, t('Create or manage a form'), 'https://forms.office.com/'),
		field('Text to display', title),
		field('Microsoft Forms response link', address),
		button,
	);
	dialog.append(form);
	return new Promise((resolve) => {
		address.addEventListener('input', () => address.setCustomValidity(''));
		form.addEventListener('submit', (event) => {
			event.preventDefault();
			const url = microsoftFormsResponseUrl(address.value);
			if (!url) {
				address.setCustomValidity(t('Enter a Microsoft Forms response link.'));
				address.reportValidity();
				return;
			}
			resolve({ title: title.value.trim() || t('Open form'), url });
			dialog.close();
		});
		dialog.addEventListener('office-dialog-close', () => resolve(undefined), { once: true });
		dialog.show();
		address.focus();
	});
}
