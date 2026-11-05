/** File > Open: this device, plus one place per cloud provider the host passed. */
import {
	defineCloudFiles,
	type OfficeCloudFileOpenEvent,
	type OfficeUiCloudFiles,
} from '../../controls';
import { cloudLabels } from './cloud-labels';
import { heading, paragraph, primary, runFile, type PageContext } from './parts';

/** Which place the Open page last showed, so a re-render (locale, sign-in) keeps it. */
const lastPlace = new WeakMap<PageContext['host'], string>();
const DEVICE = 'device';

function device(page: PageContext): HTMLElement[] {
	const { t } = page;
	return [
		paragraph(
			page,
			t('Open an Excel workbook (.xlsx, .xlsm, .xls) or a CSV file from this device.'),
		),
		primary(page, t('Browse...'), runFile(page, 'open')),
		paragraph(page, t('Files are processed entirely in the browser.'), 'xve-backstage-muted'),
	];
}

export function renderOpen(page: PageContext): void {
	const { t, host, doc } = page;
	const providers = host.fileProviders?.() ?? [];
	if (!providers.length) {
		page.content.replaceChildren(heading(page, t('Open')), ...device(page));
		return;
	}
	defineCloudFiles();
	const places = doc.createElement('div');
	places.className = 'xve-backstage-places';
	places.setAttribute('role', 'tablist');
	places.setAttribute('aria-label', t('Places'));
	const panel = doc.createElement('div');
	panel.className = 'xve-backstage-place';
	panel.setAttribute('role', 'tabpanel');
	const ids = [DEVICE, ...providers.map((provider) => provider.id)];
	const buttons = new Map<string, HTMLButtonElement>();
	const cloud = doc.createElement('office-ui-cloud-files') as OfficeUiCloudFiles;
	cloud.className = 'xve-backstage-cloud';
	cloud.accept = ['xlsx'];
	cloud.labels = cloudLabels(t);
	cloud.locale = host.options().locale;
	cloud.addEventListener('office-cloud-file-open', (event) => {
		event.stopPropagation();
		host.openCloudFile?.((event as OfficeCloudFileOpenEvent).detail.file);
	});
	const select = (id: string) => {
		const place = ids.includes(id) ? id : DEVICE;
		lastPlace.set(host, place);
		for (const [key, button] of buttons) {
			button.setAttribute('aria-selected', String(key === place));
			button.tabIndex = key === place ? 0 : -1;
		}
		if (place === DEVICE) panel.replaceChildren(...device(page));
		else {
			cloud.providers = providers.filter((provider) => provider.id === place);
			panel.replaceChildren(cloud);
		}
	};
	const label = (id: string) =>
		id === DEVICE ? t('This device') : providers.find((p) => p.id === id)!.label;
	for (const id of ids) {
		const button = doc.createElement('button');
		button.type = 'button';
		button.className = 'xve-backstage-place-tab';
		button.setAttribute('role', 'tab');
		button.dataset.place = id;
		button.textContent = label(id);
		button.addEventListener('click', () => select(id));
		button.addEventListener('keydown', (event) => {
			const step = event.key === 'ArrowRight' ? 1 : event.key === 'ArrowLeft' ? -1 : 0;
			if (!step) return;
			const next = ids[(ids.indexOf(id) + step + ids.length) % ids.length]!;
			select(next);
			buttons.get(next)?.focus();
			event.preventDefault();
		});
		buttons.set(id, button);
		places.append(button);
	}
	page.content.replaceChildren(heading(page, t('Open')), places, panel);
	select(lastPlace.get(host) ?? DEVICE);
}
