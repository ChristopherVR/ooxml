import type { OfficeConnectedServicesLabels } from './connected-services';

/**
 * The labels of `<office-ui-connected-services>` through a product's translator `t`, which maps
 * the English wording below to the display language. Every product dictionary carries these keys.
 */
export function connectedServicesLabels(
	t: (text: string) => string,
): OfficeConnectedServicesLabels {
	return {
		heading: t('Connected services'),
		description: t('Link your storage to open and save files from every app.'),
		notConnected: t('Not connected'),
		connect: t('Connect'),
		disconnect: t('Disconnect'),
		managed: t('Managed by your organization'),
		failed: t('Could not update {provider}. Try again.'),
	};
}
