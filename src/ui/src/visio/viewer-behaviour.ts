import { VISIO_LOCALES, type VisioLocale } from './ribbon-locale';
import { isScreenTipStyle, type ScreenTipStyle } from './viewer-screen-tips';

/** Undo depth is bounded by the retained drawing snapshots (see EDIT_HISTORY_LIMITS). */
export const MAX_UNDO_LEVELS = 19;
/** The AutoRecover interval, in minutes, Options accepts. */
export const MAX_RECOVER_MINUTES = 120;

/** The main tabs Customize Ribbon can hide; File is the backstage and always stays. */
export const CUSTOMIZABLE_TABS = [
	['home', 'Home'],
	['insert', 'Insert'],
	['design', 'Design'],
	['data', 'Data'],
	['process', 'Process'],
	['review', 'Review'],
	['view', 'View'],
	['help', 'Help'],
] as const;
export type CustomizableTab = (typeof CUSTOMIZABLE_TABS)[number][0];

/** The display languages, each named in its own language. */
export const DISPLAY_LANGUAGES: readonly (readonly [code: VisioLocale, name: string])[] = [
	['en', 'English'],
	['de', 'Deutsch'],
	['es', 'Español'],
	['fr', 'Français'],
	['zh-CN', '简体中文'],
];

/** The Quick Access Toolbar's commands, in the order the title bar shows them. */
export const QUICK_ACCESS_COMMANDS = [
	['save', 'Save'],
	['undo', 'Undo'],
	['redo', 'Redo'],
] as const;
export type QuickAccessCommand = (typeof QUICK_ACCESS_COMMANDS)[number][0];

/**
 * The File > Options settings that change how the viewer behaves, remembered on this device. Every
 * value is a primitive so the options dialog edits the record directly: a tab or toolbar command
 * is a boolean named `tab-<id>` or `qat-<id>`.
 */
type TabSettings = { [K in CustomizableTab as `tab-${K}`]: boolean };
type QuickAccessSettings = { [K in QuickAccessCommand as `qat-${K}`]: boolean };
export interface VisioBehaviour extends TabSettings, QuickAccessSettings {
	screenTips: ScreenTipStyle;
	undoLevels: number;
	/** Floating text formatting toolbar on a right-click and on selected text. */
	miniToolbar: boolean;
	/** Hovering a style or theme in a gallery shows it on the drawing before it is applied. */
	livePreview: boolean;
	/** The in-place text editor grows to fit text taller than its shape. */
	autoSizeText: boolean;
	/** Display language of the ribbon; empty follows the host, the page and the browser. */
	displayLanguage: VisioLocale | '';
	/** The browser's spell checker underlines mistakes in the in-place text editor. */
	checkSpelling: boolean;
	/** Keep a copy of a drawing with unsaved changes in this browser (off unless chosen). */
	autoRecover: boolean;
	/** Minutes between AutoRecover copies. */
	autoRecoverMinutes: number;
}

const tabDefaults = Object.fromEntries(
	CUSTOMIZABLE_TABS.map(([id]) => [`tab-${id}`, true]),
) as Record<`tab-${CustomizableTab}`, boolean>;
const qatDefaults = Object.fromEntries(
	QUICK_ACCESS_COMMANDS.map(([id]) => [`qat-${id}`, true]),
) as Record<`qat-${QuickAccessCommand}`, boolean>;

export const DEFAULT_BEHAVIOUR: Readonly<VisioBehaviour> = {
	screenTips: 'descriptions',
	undoLevels: MAX_UNDO_LEVELS,
	miniToolbar: true,
	livePreview: true,
	autoSizeText: true,
	displayLanguage: '',
	checkSpelling: true,
	autoRecover: false,
	autoRecoverMinutes: 10,
	...tabDefaults,
	...qatDefaults,
};

const clampUndo = (value: unknown, fallback: number): number => {
	const levels = Math.round(Number(value));
	return Number.isFinite(levels) ? Math.min(MAX_UNDO_LEVELS, Math.max(1, levels)) : fallback;
};
const clampMinutes = (value: unknown, fallback: number): number => {
	const minutes = Math.round(Number(value));
	return Number.isFinite(minutes) ? Math.min(MAX_RECOVER_MINUTES, Math.max(1, minutes)) : fallback;
};
const isLanguage = (value: unknown): value is VisioLocale | '' =>
	value === '' || (VISIO_LOCALES as readonly unknown[]).includes(value);

/**
 * `source` (a saved record, or the dialog's values) laid over `base`: a value of the wrong type is
 * ignored, a number is clamped, and nothing outside the settings above is kept.
 */
export function mergeBehaviour(
	base: Readonly<VisioBehaviour>,
	source: Readonly<Record<string, unknown>>,
): VisioBehaviour {
	const next: VisioBehaviour = { ...base };
	for (const key of Object.keys(DEFAULT_BEHAVIOUR) as (keyof VisioBehaviour)[]) {
		const value = source[key];
		if (value === undefined) continue;
		if (key === 'screenTips') {
			if (isScreenTipStyle(value)) next.screenTips = value;
		} else if (key === 'undoLevels') next.undoLevels = clampUndo(value, base.undoLevels);
		else if (key === 'autoRecoverMinutes')
			next.autoRecoverMinutes = clampMinutes(value, base.autoRecoverMinutes);
		else if (key === 'displayLanguage') {
			if (isLanguage(value)) next.displayLanguage = value;
		} else if (typeof value === 'boolean')
			(next as unknown as Record<string, unknown>)[key] = value;
	}
	// The ribbon always keeps one tab: hiding every tab shows Home.
	if (CUSTOMIZABLE_TABS.every(([id]) => !next[`tab-${id}`])) next['tab-home'] = true;
	return next;
}

const KEY = 'ooxml-visio-options';

/** The behaviour saved on this device (storage may be unavailable or hold junk). */
export function readBehaviour(): VisioBehaviour {
	try {
		const stored: unknown = JSON.parse(globalThis.localStorage?.getItem(KEY) ?? '{}');
		return mergeBehaviour(
			DEFAULT_BEHAVIOUR,
			stored && typeof stored === 'object' ? (stored as Record<string, unknown>) : {},
		);
	} catch {
		return { ...DEFAULT_BEHAVIOUR };
	}
}

export function writeBehaviour(behaviour: VisioBehaviour): void {
	try {
		globalThis.localStorage?.setItem(KEY, JSON.stringify(behaviour));
	} catch {
		// Kept for this session only.
	}
}

/** Whether two behaviours differ in any setting. */
export const behaviourChanged = (a: VisioBehaviour, b: VisioBehaviour): boolean =>
	(Object.keys(DEFAULT_BEHAVIOUR) as (keyof VisioBehaviour)[]).some((key) => a[key] !== b[key]);
