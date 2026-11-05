/** File > Open: this device, plus one place per cloud provider the host passed. */
import {
	defineOpenPlaces,
	type OfficeCloudFileOpenEvent,
	type OfficeUiOpenPlaces,
} from '../../controls';
import { cloudLabels } from './cloud-labels';
import { heading, paragraph, primary, runFile, type PageContext } from './parts';

/** Which place the Open page last showed, so a re-render (locale, sign-in) keeps it. */
const lastPlace = new WeakMap<PageContext['host'], string>();

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
	defineOpenPlaces();
	const places = doc.createElement('office-ui-open-places') as OfficeUiOpenPlaces;
	places.providers = providers;
	places.accept = ['xlsx'];
	places.labels = { ...cloudLabels(t), thisDevice: t('This device'), places: t('Places') };
	places.locale = host.options().locale;
	places.selected = lastPlace.get(host) ?? places.selected;
	places.addEventListener('office-open-place-change', (event) =>
		lastPlace.set(host, (event as CustomEvent<{ place: string }>).detail.place),
	);
	places.addEventListener('office-cloud-file-open', (event) => {
		event.stopPropagation();
		host.openCloudFile?.((event as OfficeCloudFileOpenEvent).detail.file);
	});
	places.append(...device(page));
	page.content.replaceChildren(heading(page, t('Open')), places);
}
