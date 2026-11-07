import { $, notify } from './ui.js';
import { applyOfficeTheme, applyPowerPointTheme } from './theme-tokens.js';
import { applyWorkspaceSettings, onSettings, settings } from './settings.js';

/** Each accent: its name, the light-mode value and the dark-mode value. */
export const accents = {
	blue: ['Office blue', '#0f6cbd', '#479ef5'],
	purple: ['Violet', '#5b5fc7', '#a3a5f0'],
	teal: ['Teal', '#03787c', '#4cc2c7'],
	green: ['Forest', '#107c41', '#54b054'],
	orange: ['Amber', '#a84b00', '#f7924a'],
	rose: ['Rose', '#b4126e', '#f27ac2'],
	cranberry: ['Cranberry', '#a4262c', '#f1707b'],
	graphite: ['Graphite', '#525252', '#a8a8a8'],
};
/** The accent as text on the canvas: the same in light mode, lifted for contrast in dark mode. */
const accentInk = {
	blue: ['#0f6cbd', '#7ab8ff'],
	purple: ['#5b5fc7', '#bdbef8'],
	teal: ['#03787c', '#7ad8dc'],
	green: ['#107c41', '#7ccf8a'],
	orange: ['#a84b00', '#ffae73'],
	rose: ['#b4126e', '#f7a3d6'],
	cranberry: ['#a4262c', '#f7a1a8'],
	graphite: ['#525252', '#c4c4c4'],
};
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
	return { mode, accent: accents[preferences.accent][mode === 'dark' ? 2 : 1] };
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
	root.style.setProperty(
		'--accent-ink',
		accentInk[preferences.accent][theme.mode === 'dark' ? 1 : 0],
	);
	root.style.setProperty('--office-accent', theme.accent);
	root.style.setProperty('--office-accent-hover', theme.accent);
	root.style.setProperty(
		'--office-accent-foreground',
		theme.mode === 'dark' ? '#141414' : '#ffffff',
	);
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
