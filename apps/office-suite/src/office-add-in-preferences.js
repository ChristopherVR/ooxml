import { createAddInPreferencesStore } from 'ooxml-ui/suite';

const profiles = new Map();

export function officeAddInPreferences(profile) {
	if (!profiles.has(profile)) {
		profiles.set(profile, createAddInPreferencesStore(`ooxml-suite-add-ins-${profile}`));
	}
	return profiles.get(profile);
}
