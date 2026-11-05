import { mountCloudFilesPlace, resolveOpenPlaces, selectOpenPlace } from 'ooxml-ui/pptx';
import type { CloudFileProvider, CloudFilesPlace, CloudOpenedFile } from 'ooxml-ui/pptx';

import type { Translator } from '../../../i18n';
import { createEl } from '../../../render';
import type { BackstageCloudFiles } from '../ribbon-types';
import { button } from './file-tab-dom';

export interface OpenPlacesOptions {
	cloud: BackstageCloudFiles;
	/** The place the user picked last (undefined: "This device"). */
	selectedId: string | undefined;
	onSelect(id: string): void;
	/** A file was opened in a cloud place. */
	onOpen(file: CloudOpenedFile): void;
}

export interface OpenPlacesView {
	/** The cloud provider whose place is shown, or undefined for "This device" (or no places). */
	provider: CloudFileProvider | undefined;
	/** Unmount the place's file browser. */
	destroy(): void;
}

/**
 * File > Open's places (shared `render/cloud-open`): a tablist of "This device" plus one place per
 * host provider, and the selected cloud place's `pptx-ui-cloud-files` browser. Appends nothing
 * when the host passed no providers.
 */
export function renderOpenPlaces(
	doc: Document,
	t: Translator,
	main: HTMLElement,
	options: OpenPlacesOptions,
): OpenPlacesView {
	const places = resolveOpenPlaces(options.cloud.providers, t);
	const place = selectOpenPlace(places, options.selectedId);
	if (!place) {
		return { provider: undefined, destroy: () => {} };
	}
	const tablist = createEl(doc, 'div', 'pptxv-bs-places');
	tablist.setAttribute('role', 'tablist');
	tablist.setAttribute('aria-label', t('pptx.backstage.places'));
	for (const item of places) {
		const selected = item.id === place.id;
		const tab = button(doc, item.label, () => options.onSelect(item.id));
		tab.setAttribute('role', 'tab');
		tab.setAttribute('data-pptx-open-place', item.id);
		tab.setAttribute('aria-selected', String(selected));
		tab.classList.toggle('active', selected);
		tablist.appendChild(tab);
	}
	main.appendChild(tablist);
	const provider = place.provider;
	if (!provider) {
		return { provider: undefined, destroy: () => {} };
	}
	const host = createEl(doc, 'div', 'pptxv-bs-cloud-place');
	host.setAttribute('data-pptx-cloud-place', provider.id);
	main.appendChild(host);
	const mounted: CloudFilesPlace = mountCloudFilesPlace(host, {
		provider,
		translate: t,
		locale: options.cloud.locale,
		onOpen: options.onOpen,
	});
	return { provider, destroy: () => mounted.destroy() };
}
