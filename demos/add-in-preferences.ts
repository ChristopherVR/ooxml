import { createAddInPreferencesStore } from '../src/ui/src/suite/add-in-preferences-store';

export function demoAddInPreferences() {
	if (new URLSearchParams(location.search).get('addInsStorage') === 'off') return null;
	return createAddInPreferencesStore('ooxml-demo-add-ins');
}
