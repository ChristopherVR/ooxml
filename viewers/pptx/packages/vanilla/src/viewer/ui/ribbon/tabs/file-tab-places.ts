import { CLOUD_OPEN_ACCEPT, openPlacesLabels } from 'ooxml-ui/pptx';
import type { CloudOpenedFile, OfficeOpenPlacesLabels } from 'ooxml-ui/pptx';

import type { Translator } from '../../../i18n';
import type { BackstageCloudFiles } from '../ribbon-types';

export interface OpenPlacesOptions {
	cloud: BackstageCloudFiles;
	/** The place the user picked last (undefined: "This device"). */
	selectedId: string | undefined;
	onSelect(id: string): void;
	/** A file was opened in a cloud place. */
	onOpen(file: CloudOpenedFile): void;
}

/** The shared places element's properties the binding sets. */
interface OpenPlacesElement extends HTMLElement {
	providers: BackstageCloudFiles['providers'];
	accept: readonly string[];
	labels: OfficeOpenPlacesLabels;
	locale: string | undefined;
	selected: string;
}

/**
 * The container of File > Open's device content (search, Browse, recent files): the shared
 * `<pptx-ui-open-places>`, which adds a tab per host provider (none without `fileProviders`)
 * and shows the provider's file browser for the chosen one. Home passes no providers, so it is
 * a plain wrapper there.
 */
export function createOpenPlaces(
	doc: Document,
	t: Translator,
	options: OpenPlacesOptions,
): HTMLElement {
	const element = doc.createElement('pptx-ui-open-places') as OpenPlacesElement;
	element.className = 'pptxv-bs-places';
	element.providers = options.cloud.providers;
	element.accept = CLOUD_OPEN_ACCEPT;
	element.labels = openPlacesLabels(t);
	element.locale = options.cloud.locale;
	element.selected = options.selectedId ?? 'device';
	element.addEventListener('office-open-place-change', (event) =>
		options.onSelect((event as CustomEvent<{ place: string }>).detail.place),
	);
	element.addEventListener('office-cloud-file-open', (event) =>
		options.onOpen((event as CustomEvent<{ file: CloudOpenedFile }>).detail.file),
	);
	return element;
}
