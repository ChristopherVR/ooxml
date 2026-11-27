import { createAddInPreferencesStore, setOfficeAddInPreferences } from 'ooxml-ui/suite';

import { demoAddInPreferences } from '../../add-in-preferences';

export function enableDemoAddInPreferences(host: HTMLElement) {
	setOfficeAddInPreferences(host, demoAddInPreferences(createAddInPreferencesStore));
}
