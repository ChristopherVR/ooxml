/**
 * The suite's own preferences that every editor it opens honours: who you are when you comment,
 * the display language, which parts of each editor's chrome show, the corner style and motion of
 * the workspace, and the privacy switches the editors expose. Appearance (`themes.js`, key
 * `ooxml-suite-theme`, read before first paint) and font downloads (`font-downloads.js`, key
 * `ooxml-suite-fonts`) keep their own keys; this module owns `ooxml-suite-settings`.
 *
 * A value left at "each app's own choice" (`locale: 'auto'`, `blockExternalContent: false`) passes
 * nothing to the editors, so the choice each one keeps in its own File > Options still applies.
 */
export const SETTINGS_KEY = 'ooxml-suite-settings';

/** The display languages every editor ships (Word, Excel, PowerPoint and Visio share them). */
export const LOCALES = [
	['en', 'English'],
	['fr', 'Français'],
	['es', 'Español'],
	['de', 'Deutsch'],
	['zh-CN', '简体中文'],
];

/** Corner styles: the workspace's control and panel radii, and the editors' radius. */
export const CORNERS = {
	square: { control: 0, panel: 0, label: 'Square' },
	rounded: { control: 4, panel: 8, label: 'Rounded' },
	round: { control: 8, panel: 14, label: 'Round' },
};

export const DEFAULT_SETTINGS = Object.freeze({
	authorName: '',
	locale: 'auto',
	showRibbon: true,
	showFormulaBar: true,
	showThumbnails: true,
	measurementUnit: 'inches',
	blockExternalContent: false,
	corners: 'rounded',
	reduceMotion: false,
	confirmClose: true,
});

const choices = {
	locale: ['auto', ...LOCALES.map(([code]) => code)],
	measurementUnit: ['inches', 'centimetres'],
	corners: Object.keys(CORNERS),
};

/** Every known key with a value of the right type and range; unknown keys are dropped. */
export function normalizeSettings(raw) {
	const next = { ...DEFAULT_SETTINGS };
	if (!raw || typeof raw !== 'object') return next;
	for (const [key, fallback] of Object.entries(DEFAULT_SETTINGS)) {
		const value = raw[key];
		if (typeof fallback === 'boolean') {
			if (typeof value === 'boolean') next[key] = value;
		} else if (choices[key]) {
			if (choices[key].includes(value)) next[key] = value;
		} else if (typeof value === 'string') {
			// Free text (the author name): one line, trimmed, bounded.
			next[key] = value
				.replace(/[\r\n\t]+/g, ' ')
				.trim()
				.slice(0, 80);
		}
	}
	return next;
}

function read(raw) {
	try {
		return normalizeSettings(JSON.parse(raw));
	} catch {
		return normalizeSettings(null);
	}
}

let current = normalizeSettings(null);
try {
	current = read(globalThis.localStorage?.getItem(SETTINGS_KEY) ?? 'null');
} catch {}
const listeners = new Set();

export const settings = () => ({ ...current });

function publish(previous) {
	const changed = Object.keys(current).filter((key) => current[key] !== previous[key]);
	if (!changed.length) return;
	for (const listener of listeners) listener(settings(), changed);
}

/** Merge `patch`, save it and tell every listener which keys changed. */
export function updateSettings(patch) {
	const previous = current;
	current = normalizeSettings({ ...current, ...patch });
	try {
		globalThis.localStorage?.setItem(SETTINGS_KEY, JSON.stringify(current));
	} catch {}
	publish(previous);
	return settings();
}

export function resetSettings() {
	return updateSettings(DEFAULT_SETTINGS);
}

/** `listener(settings, changedKeys)` on every change, here or in another tab. */
export function onSettings(listener) {
	listeners.add(listener);
	return () => listeners.delete(listener);
}

/** The locale to pass to the editors, or `undefined` to keep each one's own choice. */
export const editorLocale = (s = current) => (s.locale === 'auto' ? undefined : s.locale);

/** The author name for comments and tracked changes: the setting, else the profile's name. */
export const authorName = (profileName, s = current) =>
	s.authorName || (typeof profileName === 'string' ? profileName.trim() : '') || undefined;

/** The editors' corner radius for the corner style. */
export const editorRadius = (s = current) => `${(CORNERS[s.corners] ?? CORNERS.rounded).control}px`;

/** PowerPoint's `allowExternalContent`: `false` when blocked, else `undefined` (its own choice). */
export const externalContentOption = (s = current) => (s.blockExternalContent ? false : undefined);

/** The workspace's corner style and motion, on the root element. */
export function applyWorkspaceSettings(root = globalThis.document?.documentElement, s = current) {
	if (!root) return;
	const corners = CORNERS[s.corners] ?? CORNERS.rounded;
	root.style.setProperty('--radius', `${corners.control}px`);
	root.style.setProperty('--radius-panel', `${corners.panel}px`);
	root.style.setProperty('--office-radius', `${corners.control}px`);
	root.classList.toggle('reduce-motion', s.reduceMotion);
	if (s.reduceMotion) root.classList.remove('theme-anim');
}

// Another tab of the suite changed a setting.
globalThis.addEventListener?.('storage', (event) => {
	if (event.key !== SETTINGS_KEY) return;
	const previous = current;
	current = read(event.newValue);
	publish(previous);
});
