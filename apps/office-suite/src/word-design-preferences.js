import { currentProfile } from './profile-state.js';
import { settings } from './settings.js';

export function wordDesignPreferences() {
	if (!settings().rememberWordDesign) return undefined;
	const key = `ooxml-suite-word-design-${currentProfile().id}`;
	return {
		read() {
			try {
				const raw = localStorage.getItem(key);
				return raw && raw.length <= 8192 ? JSON.parse(raw) : undefined;
			} catch {
				return undefined;
			}
		},
		write(value) {
			localStorage.setItem(key, JSON.stringify(value));
		},
	};
}
