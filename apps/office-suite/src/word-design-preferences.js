import { currentProfile } from './profile-state.js';
import { settings } from './settings.js';

export function wordDesignPreferences() {
	if (!settings().rememberWordDesign) return undefined;
	const key = `ooxml-suite-word-design-${currentProfile().id}`;
	const store = (name, limit) => ({
		read() {
			try {
				const raw = localStorage.getItem(name);
				return raw && raw.length <= limit ? JSON.parse(raw) : undefined;
			} catch {
				return undefined;
			}
		},
		write(value) {
			localStorage.setItem(name, JSON.stringify(value));
		},
	});
	return { ...store(key, 8192), schemes: store(`${key}-schemes`, 48 * 1024) };
}
