import type { CloudFileProvider, CloudOpenedFile } from 'ooxml-core/embedding';
import { defineOpenPlaces } from '../controls';
import type { OfficeCloudFileOpenEvent, OfficeUiOpenPlaces } from '../controls';
import { cloudFileLabels } from './cloud-labels';
import type { PageContext } from './backstage-pages';

/** The places File > Open lists: this device first, then one per cloud provider. */
export function renderOpenPlaces(
	{ handlers, t, content }: PageContext,
	providers: readonly CloudFileProvider[],
	device: () => HTMLElement[],
	heading: HTMLElement,
): void {
	defineOpenPlaces();
	const places = document.createElement('office-ui-open-places') as OfficeUiOpenPlaces;
	places.providers = providers;
	places.accept = ['docx'];
	places.labels = {
		...cloudFileLabels(handlers.locale()),
		thisDevice: t('This device'),
		places: t('cloud.providers'),
	};
	places.locale = handlers.locale();
	places.addEventListener('office-cloud-file-open', (event) => {
		const { file } = (event as OfficeCloudFileOpenEvent).detail;
		void handlers.openCloudFile(file satisfies CloudOpenedFile);
	});
	places.append(...device());
	content.replaceChildren(heading, places);
}
