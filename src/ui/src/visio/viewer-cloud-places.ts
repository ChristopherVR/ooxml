import type { CloudFileProvider, CloudOpenedFile } from 'ooxml-core/embedding';
import { defineOpenPlaces } from '../controls';
import type { OfficeUiOpenPlaces } from '../files/open-places';
import { VISIO_CLOUD_FILES_LABELS } from './locales/cloud-files';
import { translateRibbonText, type VisioLocale } from './ribbon-locale';

/**
 * The places of File > Open when the host passes `fileProviders`: the shared
 * `<office-ui-open-places>` already on the page wraps This device (the Browse tile) and shows
 * the cloud file browser for `.vsdx` drawings per provider. Without providers it shows only
 * the Browse tile.
 */
export class CloudPlaces {
	readonly #element: OfficeUiOpenPlaces;
	#providers: readonly CloudFileProvider[] = [];

	constructor(
		section: HTMLElement,
		private readonly onOpen: (file: CloudOpenedFile, providerId: string) => void,
		signal: AbortSignal,
	) {
		defineOpenPlaces();
		this.#element = section.querySelector<OfficeUiOpenPlaces>('office-ui-open-places')!;
		this.#element.accept = ['vsdx'];
		this.#element.addEventListener(
			'office-cloud-file-open',
			(event) => {
				const { providerId, file } = (
					event as CustomEvent<{ providerId: string; file: CloudOpenedFile }>
				).detail;
				this.onOpen(file, providerId);
			},
			{ signal },
		);
	}

	get providers(): readonly CloudFileProvider[] {
		return this.#providers;
	}

	setProviders(providers: readonly CloudFileProvider[] | null | undefined): void {
		this.#providers = [...(providers ?? [])];
		this.#element.providers = this.#providers;
	}

	setLocale(locale: VisioLocale): void {
		this.#element.locale = locale;
		this.#element.labels = {
			...VISIO_CLOUD_FILES_LABELS[locale],
			thisDevice: translateRibbonText(locale, 'This device'),
			places: translateRibbonText(locale, 'Places'),
		};
	}
}
