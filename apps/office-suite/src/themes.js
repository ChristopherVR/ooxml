import { $, notify } from './ui.js';
import { accents, applyOfficeTheme, applyPowerPointTheme, resolveTheme } from './theme-tokens.js';
import { applyWorkspaceSettings, onSettings, settings } from './settings.js';

export const DEFAULT_THEME = Object.freeze({ mode: 'dark', accent: 'blue' });
/** A first visit opens in dark mode; the appearance settings remember any other choice. */
let preferences = { ...DEFAULT_THEME };
try {
	const saved = JSON.parse(localStorage.getItem('ooxml-suite-theme'));
	if (saved && ['light', 'dark', 'system'].includes(saved.mode) && accents[saved.accent])
		preferences = saved;
} catch {}
const media = matchMedia('(prefers-color-scheme: dark)');
const listeners = new Set();
export function currentTheme() {
	const mode =
		preferences.mode === 'system' ? (media.matches ? 'dark' : 'light') : preferences.mode;
	return resolveTheme(mode, preferences.accent);
}
export function onTheme(listener) {
	listeners.add(listener);
	listener(currentTheme());
	return () => listeners.delete(listener);
}
function apply() {
	const theme = currentTheme(),
		root = document.documentElement;
	applyOfficeTheme(root, theme);
	applyPowerPointTheme(root, theme);
	root.dataset.theme = theme.mode;
	root.dataset.officeTheme = theme.mode;
	root.style.setProperty('--blue', theme.accent);
	root.style.setProperty('--accent-ink', theme.accentInk);
	root.style.colorScheme = theme.mode;
	for (const listener of listeners) listener(theme);
}
media.addEventListener('change', apply);
/** The saved choice: `mode` ('light', 'dark' or 'system') and `accent` (a key of `accents`). */
export const themePreferences = () => ({ ...preferences });
export function setThemePreferences(next) {
	const mode = ['system', 'light', 'dark'].includes(next?.mode) ? next.mode : preferences.mode;
	const accent = accents[next?.accent] ? next.accent : preferences.accent;
	if (mode === preferences.mode && accent === preferences.accent) return;
	preferences = { mode, accent };
	try {
		localStorage.setItem('ooxml-suite-theme', JSON.stringify(preferences));
	} catch {}
	apply();
}
export function setThemeMode(mode) {
	setThemePreferences({ mode });
}
window.addEventListener('storage', (event) => {
	if (event.key !== 'ooxml-suite-theme') return;
	try {
		const saved = JSON.parse(event.newValue);
		if (saved && ['light', 'dark', 'system'].includes(saved.mode) && accents[saved.accent]) {
			preferences = saved;
			apply();
		}
	} catch {}
});
export function mountThemes() {
	apply();
	applyWorkspaceSettings();
	onSettings((next, changed) => {
		if (changed.includes('locale')) applyWorkspaceSettings(document.documentElement, next);
		if (changed.some((key) => key === 'corners' || key === 'reduceMotion')) {
			applyWorkspaceSettings(document.documentElement, next);
			if (!next.reduceMotion) document.documentElement.classList.add('theme-anim');
			// Editors take their corner radius with the palette.
			for (const listener of listeners) listener(currentTheme());
		}
	});
	// Palette changes after the first paint cross-fade (styles.css); the first paint is instant.
	requestAnimationFrame(() => {
		if (!settings().reduceMotion) document.documentElement.classList.add('theme-anim');
	});
	$('theme-picker').onclick = () =>
		void import('./settings-dialog.js').then(({ openSettings }) => openSettings()).catch(notify);
}
