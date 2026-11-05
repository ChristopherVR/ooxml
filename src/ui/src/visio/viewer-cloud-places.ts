import type { CloudFileProvider, CloudOpenedFile } from 'ooxml-core/embedding';
import { defineCloudFiles } from '../controls';
import type { OfficeUiCloudFiles } from '../files/cloud-files';
import { VISIO_CLOUD_FILES_LABELS } from './locales/cloud-files';
import type { VisioLocale } from './ribbon-locale';

const DEVICE = 'device';

/**
 * The places of File > Open when the host passes `fileProviders`: This device (the Browse tile
 * already on the page) and one place per provider, which shows the shared cloud file browser for
 * `.vsdx` drawings. Without providers the page is untouched.
 */
export class CloudPlaces {
	readonly #section: HTMLElement;
	readonly #places: HTMLElement;
	readonly #cloud: HTMLElement;
	#browser: OfficeUiCloudFiles | undefined;
	#providers: readonly CloudFileProvider[] = [];
	#selected = DEVICE;
	#locale: VisioLocale = 'en';

	constructor(
		section: HTMLElement,
		private readonly onOpen: (file: CloudOpenedFile, providerId: string) => void,
		signal: AbortSignal,
	) {
		this.#section = section;
		this.#places = section.querySelector<HTMLElement>('.backstage-places')!;
		this.#cloud = section.querySelector<HTMLElement>('.backstage-cloud')!;
		this.#places.addEventListener(
			'click',
			(event) => {
				const button = (event.target as Element).closest<HTMLButtonElement>(
					'[data-backstage-place]',
				);
				if (button) this.select(button.dataset.backstagePlace!);
			},
			{ signal },
		);
		this.#cloud.addEventListener(
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
		if (this.#selected !== DEVICE && !this.#providers.some((p) => p.id === this.#selected))
			this.#selected = DEVICE;
		this.#render();
	}

	setLocale(locale: VisioLocale): void {
		this.#locale = locale;
		if (this.#browser) {
			this.#browser.locale = locale;
			this.#browser.labels = VISIO_CLOUD_FILES_LABELS[locale];
		}
	}

	select(id: string): void {
		this.#selected = id !== DEVICE && this.#providers.some((p) => p.id === id) ? id : DEVICE;
		this.#render();
	}

	#render(): void {
		const doc = this.#section.ownerDocument;
		const any = this.#providers.length > 0;
		this.#places.hidden = !any;
		this.#places.replaceChildren();
		if (any) {
			const entries: Array<[string, string, boolean]> = [
				[DEVICE, 'This device', false],
				...this.#providers.map((p): [string, string, boolean] => [p.id, p.label, true]),
			];
			for (const [id, label, user] of entries) {
				const button = doc.createElement('button');
				button.type = 'button';
				button.className = 'backstage-place';
				button.dataset.backstagePlace = id;
				button.setAttribute('aria-pressed', String(id === this.#selected));
				button.textContent = label;
				if (user) button.dataset.userText = '';
				this.#places.append(button);
			}
		}
		const cloud = any && this.#selected !== DEVICE;
		for (const node of this.#section.querySelectorAll<HTMLElement>('[data-place="device"]'))
			node.hidden = cloud;
		this.#cloud.hidden = !cloud;
		const provider = this.#providers.find((p) => p.id === this.#selected);
		if (!cloud || !provider) {
			this.#browser?.remove();
			this.#browser = undefined;
			return;
		}
		if (!this.#browser) {
			defineCloudFiles();
			this.#browser = doc.createElement('office-ui-cloud-files') as OfficeUiCloudFiles;
			this.#browser.accept = ['vsdx'];
			this.#cloud.append(this.#browser);
			this.setLocale(this.#locale);
		}
		this.#browser.providers = [provider];
		this.#browser.active = provider.id;
	}
}
