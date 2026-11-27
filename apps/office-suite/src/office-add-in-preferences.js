import { createAddInPreferencesStore } from '../../../src/ui/src/suite/add-in-preferences-store.ts';

const profiles = new Map();

export function officeAddInPreferences(profile) {
	if (!profiles.has(profile)) {
		profiles.set(profile, createAddInPreferencesStore(`ooxml-suite-add-ins-${profile}`));
	}
	return profiles.get(profile);
}
