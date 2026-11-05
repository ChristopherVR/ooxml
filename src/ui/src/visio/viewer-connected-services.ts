import type { CloudFileProvider } from 'ooxml-core/embedding';
import { defineConnectedServices } from '../controls';
import type { OfficeUiConnectedServices } from '../files/connected-services';
import { connectedServicesLabels } from '../files/connected-services-labels';
import { translateRibbonText, type VisioLocale } from './ribbon-locale';

/**
 * The Account page's "Connected services" list when the host passes `fileProviders`: the shared
 * `<office-ui-connected-services>` already on the page, handed the same providers File > Open
 * uses. Without providers it renders nothing.
 */
export class ConnectedServices {
	readonly #element: OfficeUiConnectedServices | null;

	constructor(root: ParentNode) {
		defineConnectedServices();
		this.#element = root.querySelector<OfficeUiConnectedServices>('office-ui-connected-services');
	}

	setProviders(providers: readonly CloudFileProvider[]): void {
		if (this.#element) this.#element.providers = providers;
	}

	setLocale(locale: VisioLocale): void {
		if (!this.#element) return;
		this.#element.labels = connectedServicesLabels((text) => translateRibbonText(locale, text));
	}
}
