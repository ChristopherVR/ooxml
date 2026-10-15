import { $, choose } from './ui.js';
import { applyOfficeTheme, applyPowerPointTheme } from './theme-tokens.js';

/** Each accent: its name, the light-mode value and the dark-mode value. */
const accents = {
	blue: ['Office blue', '#0f6cbd', '#479ef5'],
	purple: ['Violet', '#5b5fc7', '#a3a5f0'],
	green: ['Forest', '#107c41', '#54b054'],
	graphite: ['Graphite', '#525252', '#a8a8a8'],
};
/** The accent as text on the canvas: the same in light mode, lifted for contrast in dark mode. */
const accentInk = {
	blue: ['#0f6cbd', '#7ab8ff'],
	purple: ['#5b5fc7', '#bdbef8'],
	green: ['#107c41', '#7ccf8a'],
	graphite: ['#525252', '#c4c4c4'],
};
/** A first visit opens in dark mode; the appearance settings remember any other choice. */
let preferences = { mode: 'dark', accent: 'blue' };
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
export function setThemeMode(mode) {
	if (!['system', 'light', 'dark'].includes(mode)) return;
	preferences = { ...preferences, mode };
	try {
		localStorage.setItem('ooxml-suite-theme', JSON.stringify(preferences));
	} catch {}
	apply();
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
	// Palette changes after the first paint cross-fade (styles.css); the first paint is instant.
	requestAnimationFrame(() => document.documentElement.classList.add('theme-anim'));
	$('theme-picker').onclick = () => {
		const form = document.createElement('form');
		form.innerHTML = `<p>Choose the appearance of your workspace and all open apps.</p><fieldset><legend>Appearance</legend><div class="theme-modes">${['light', 'dark', 'system'].map((mode) => `<label><input type="radio" name="mode" value="${mode}" ${preferences.mode === mode ? 'checked' : ''} /><span class="theme-preview ${mode}"></span>${mode[0].toUpperCase() + mode.slice(1)}</label>`).join('')}</div></fieldset><fieldset><legend>Accent theme</legend><div class="theme-accents">${Object.entries(
			accents,
		)
			.map(
				([id, [name, color]]) =>
					`<label><input type="radio" name="accent" value="${id}" ${preferences.accent === id ? 'checked' : ''} /><span style="background:${color}"></span>${name}</label>`,
			)
			.join(
				'',
			)}</div></fieldset><p class="theme-hint">Documents keep their own colours. The appearance changes the workspace around them.</p><button class="primary">Done</button>`;
		form.onchange = () => {
			const data = new FormData(form);
			preferences = { mode: data.get('mode'), accent: data.get('accent') };
			try {
				localStorage.setItem('ooxml-suite-theme', JSON.stringify(preferences));
			} catch {}
			apply();
		};
		form.onsubmit = (event) => {
			event.preventDefault();
			$('dialog').close();
		};
		choose('Appearance', form);
	};
}
