/**
 * The few assistant settings a user may keep for themselves, and the settings a request actually
 * uses: the policy decides, and a personal choice only applies where `userControls` allows it.
 */
import { currentProfile } from '../profile-state.js';
import { currentPolicy } from './policy.js';

const key = () => `ooxml-suite-ai-prefs-v1-${currentProfile().id}`;
const listeners = new Set();

export function personalPreferences() {
	try {
		return JSON.parse(localStorage.getItem(key())) ?? {};
	} catch {
		return {};
	}
}
export function savePersonalPreferences(values) {
	try {
		localStorage.setItem(key(), JSON.stringify(values));
	} catch {}
	for (const listener of listeners) listener();
}
export function onPreferences(listener) {
	listeners.add(listener);
	return () => listeners.delete(listener);
}

/** The settings for the next request. */
export function effectiveSettings() {
	const { behaviour, userControls, limits } = currentPolicy();
	const mine = personalPreferences();
	const pick = (name) => (userControls[name] && mine[name] !== undefined ? mine[name] : undefined);
	return {
		style: pick('style') ?? behaviour.style,
		language: pick('language') ?? behaviour.language,
		/** The organisation's instructions always apply; a user's own are added after them. */
		organisationInstructions: behaviour.instructions,
		instructions: pick('instructions') ?? '',
		temperature: pick('temperature') ?? behaviour.temperature,
		maxTokens: limits.maxOutputTokens,
		historyMessages: limits.historyMessages,
		maxContextChars: limits.maxContextChars,
		attachOpenDocument: mine.attachOpenDocument !== false,
		openCreatedFiles: mine.openCreatedFiles !== false,
		animate: mine.animate !== false,
	};
}
