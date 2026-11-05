/** File > Account: the storage linked through the host's `fileProviders` ("Connected services"). */
import { defineConnectedServices, type OfficeUiConnectedServices } from '../../controls';
import { connectedServicesLabels } from '../../files/connected-services-labels';
import { heading, type PageContext } from './parts';

export function renderAccount(page: PageContext): void {
	const { t, host, doc } = page;
	defineConnectedServices();
	const services = doc.createElement('office-ui-connected-services') as OfficeUiConnectedServices;
	services.providers = host.fileProviders?.() ?? [];
	services.labels = connectedServicesLabels(t);
	page.content.replaceChildren(heading(page, t('Account')), services);
}
