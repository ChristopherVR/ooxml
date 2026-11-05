import { defineConnectedServices } from '../controls';
import type { OfficeUiConnectedServices } from '../controls';
import { connectedServicesLabels } from '../files/connected-services-labels';
import type { PageContext } from './backstage-pages';

/** File > Account: the storage linked through the host's `fileProviders` ("Connected services"). */
export function renderAccount({ handlers, t, content }: PageContext): void {
	defineConnectedServices();
	const title = document.createElement('h2');
	title.textContent = t('Account');
	const services = document.createElement(
		'office-ui-connected-services',
	) as OfficeUiConnectedServices;
	services.providers = handlers.fileProviders();
	services.labels = connectedServicesLabels(t);
	content.replaceChildren(title, services);
}
