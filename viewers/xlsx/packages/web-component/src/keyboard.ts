/**
 * Global keyboard shortcuts. The map only decides which command id (or shell action) a key runs;
 * the behaviour lives in the commands. Grid navigation and in-cell editing keys belong to the grid,
 * which handles them first (a handled event is `defaultPrevented` and left alone here).
 */
import type { EditorContext } from './context';

/** `Mod` is Cmd on Apple platforms and Ctrl elsewhere; `code:` matches `KeyboardEvent.code`. */
export type KeySpec = string;

export type ShellAction = 'help' | 'keytips' | 'next-region' | 'previous-region' | 'context-menu';

export interface Shortcut {
	keys: readonly KeySpec[];
	label: string;
	command?: string;
	arg?: unknown;
	action?: ShellAction;
	/** True: also works while typing in a text field (formula bar, dialogs). */
	global?: boolean;
}

/** Excel's shortcuts, mapped to command ids (`file.*` and `sheet.next/previous` are the shell's). */
export const SHORTCUTS: readonly Shortcut[] = [
	{ keys: ['Mod+S'], label: 'Save', command: 'file.save', global: true },
	{ keys: ['Mod+O'], label: 'Open', command: 'file.open', global: true },
	{ keys: ['Mod+N'], label: 'New workbook', command: 'file.new', global: true },
	{ keys: ['Mod+P'], label: 'Print', command: 'file.print', global: true },
	{ keys: ['Mod+F'], label: 'Find', command: 'home.find' },
	{ keys: ['Mod+H'], label: 'Replace', command: 'home.replace' },
	{ keys: ['Mod+G', 'F5'], label: 'Go To', command: 'home.go-to' },
	{ keys: ['Mod+Z'], label: 'Undo', command: 'edit.undo' },
	{ keys: ['Mod+Y', 'Mod+Shift+Z'], label: 'Redo', command: 'edit.redo' },
	{ keys: ['Mod+1'], label: 'Format Cells', command: 'format.cells' },
	{ keys: ['Mod+B', 'Mod+2'], label: 'Bold', command: 'home.bold' },
	{ keys: ['Mod+I', 'Mod+3'], label: 'Italic', command: 'home.italic' },
	{ keys: ['Mod+U', 'Mod+4'], label: 'Underline', command: 'home.underline' },
	{ keys: ['Mod+5'], label: 'Strikethrough', command: 'home.strikethrough' },
	{
		keys: ['Mod+Shift+~', 'Mod+Shift+code:Backquote'],
		label: 'General number format',
		command: 'home.format-general',
	},
	{
		keys: ['Mod+Shift+$', 'Mod+Shift+code:Digit4'],
		label: 'Currency format',
		command: 'home.format-currency',
	},
	{
		keys: ['Mod+Shift+%', 'Mod+Shift+code:Digit5'],
		label: 'Percentage format',
		command: 'home.format-percent',
	},
	{
		keys: ['Mod+Shift+^', 'Mod+Shift+code:Digit6'],
		label: 'Scientific format',
		command: 'home.format-scientific',
	},
	{
		keys: ['Mod+Shift+#', 'Mod+Shift+code:Digit3'],
		label: 'Date format',
		command: 'home.format-date',
	},
	{
		keys: ['Mod+Shift+@', 'Mod+Shift+code:Digit2'],
		label: 'Time format',
		command: 'home.format-time',
	},
	{
		keys: ['Mod+Shift+!', 'Mod+Shift+code:Digit1'],
		label: 'Number format',
		command: 'home.format-number',
	},
	{ keys: ['Mod+;'], label: 'Insert the current date', command: 'edit.insert-date' },
	{
		keys: ['Mod+Shift+:', 'Mod+Shift+;'],
		label: 'Insert the current time',
		command: 'edit.insert-time',
	},
	{
		keys: ['Mod+`', 'Mod+code:Backquote'],
		label: 'Show Formulas',
		command: 'formulas.show-formulas',
	},
	{ keys: ['Alt+='], label: 'AutoSum', command: 'home.autosum' },
	{ keys: ['Mod+ '], label: 'Select the entire column', command: 'edit.select-column' },
	{ keys: ['Shift+ '], label: 'Select the entire row', command: 'edit.select-row' },
	{ keys: ['Mod+PageUp'], label: 'Previous sheet', command: 'sheet.previous' },
	{ keys: ['Mod+PageDown'], label: 'Next sheet', command: 'sheet.next' },
	{ keys: ['Mod+Shift+L'], label: 'Filter', command: 'data.filter' },
	{ keys: ['Mod+T', 'Mod+L'], label: 'Create a table', command: 'insert.table' },
	{ keys: ['Mod+K'], label: 'Insert link', command: 'insert.link' },
	{ keys: ['Shift+F2'], label: 'New comment', command: 'review.new-comment' },
	{ keys: ['Mod+-'], label: 'Delete cells', command: 'cells.delete' },
	{ keys: ['Mod+Shift++', 'Mod++', 'Mod+Shift+='], label: 'Insert cells', command: 'cells.insert' },
	{ keys: ['F9'], label: 'Calculate Now', command: 'formulas.calculate-now' },
	{ keys: ['F6'], label: 'Move to the next region', action: 'next-region', global: true },
	{
		keys: ['Shift+F6'],
		label: 'Move to the previous region',
		action: 'previous-region',
		global: true,
	},
	{ keys: ['F10', 'Alt'], label: 'Show KeyTips', action: 'keytips', global: true },
	{ keys: ['Shift+F10', 'ContextMenu'], label: 'Open the context menu', action: 'context-menu' },
	{ keys: ['Mod+/', 'F1'], label: 'Keyboard shortcuts', action: 'help', global: true },
];

export interface KeyEventLike {
	key: string;
	code?: string;
	ctrlKey: boolean;
	metaKey: boolean;
	altKey: boolean;
	shiftKey: boolean;
}

interface Parsed {
	mod: boolean;
	shift: boolean;
	alt: boolean;
	key: string;
	code: string | undefined;
	tap: boolean;
}

function parse(spec: KeySpec): Parsed {
	const parts = spec.split('+');
	// `Mod++` and `Mod+ ` end in the key itself.
	let key = parts.pop() ?? '';
	if (key === '' && spec.endsWith('+')) key = '+';
	const modifiers = new Set(parts.map((part) => part.toLowerCase()).filter(Boolean));
	const code = key.startsWith('code:') ? key.slice(5) : undefined;
	return {
		mod: modifiers.has('mod'),
		shift: modifiers.has('shift'),
		alt: modifiers.has('alt'),
		key: code ? '' : key.toLowerCase(),
		code,
		tap: key === 'Alt' && !modifiers.size,
	};
}

export function isMacPlatform(): boolean {
	if (typeof navigator === 'undefined') return false;
	const data = (navigator as Navigator & { userAgentData?: { platform?: string } }).userAgentData;
	return /mac|iphone|ipad|ipod/i.test(data?.platform ?? navigator.platform ?? '');
}

/** Exact modifier match: Ctrl on a Mac, or Cmd elsewhere, never satisfies `Mod`. */
export function matchesKeys(event: KeyEventLike, spec: KeySpec, mac: boolean): boolean {
	const wanted = parse(spec);
	if (wanted.tap) return false;
	const mod = mac ? event.metaKey : event.ctrlKey;
	const foreign = mac ? event.ctrlKey : event.metaKey;
	if (
		mod !== wanted.mod ||
		foreign ||
		event.altKey !== wanted.alt ||
		event.shiftKey !== wanted.shift
	)
		return false;
	return wanted.code ? event.code === wanted.code : event.key.toLowerCase() === wanted.key;
}

const DISPLAY: Record<string, string> = {
	' ': 'Space',
	escape: 'Esc',
	contextmenu: 'Menu',
	pageup: 'PgUp',
	pagedown: 'PgDn',
};

/** Human-readable key text (`Ctrl+Shift+L`, `Cmd+S` on a Mac). Legends are not translated. */
export function formatKeys(spec: KeySpec, mac: boolean): string {
	const parsed = parse(spec);
	const key = parsed.code ?? (spec.endsWith('+') ? '+' : (spec.split('+').pop() ?? ''));
	const names = [
		parsed.mod ? (mac ? 'Cmd' : 'Ctrl') : '',
		parsed.alt ? (mac ? 'Option' : 'Alt') : '',
		parsed.shift ? 'Shift' : '',
		DISPLAY[key.toLowerCase()] ?? (key.length === 1 ? key.toUpperCase() : key),
	];
	return names.filter(Boolean).join('+');
}

/** Specs listed in help: the first of each shortcut without `code:` duplicates. */
export const displayKeys = (shortcut: Shortcut, mac: boolean): string[] =>
	shortcut.keys.filter((spec) => !spec.includes('code:')).map((spec) => formatKeys(spec, mac));

export function findShortcut(event: KeyEventLike, mac: boolean): Shortcut | undefined {
	return SHORTCUTS.find((shortcut) => shortcut.keys.some((spec) => matchesKeys(event, spec, mac)));
}

/** True while focus is in a text field (formula bar, name box, dialog inputs, Tell me). */
function inTextField(event: Event): boolean {
	const target = event.composedPath()[0];
	if (!(target instanceof Element)) return false;
	// The grid keeps focus in a hidden input even in Ready mode; it reports editing itself.
	if (target.closest('.xve-grid-host')) return false;
	return target.matches(
		'input:not([type="checkbox"]):not([type="radio"]):not([type="range"]), textarea, select, [contenteditable=""], [contenteditable="true"]',
	);
}

/**
 * Listens on the host (shadow-tree events bubble to it, retargeted) and runs the map. Returns a
 * disposer.
 */
export function installKeyboard(
	ctx: EditorContext,
	actions: (action: ShellAction) => void,
): () => void {
	const host = ctx.host;
	const mac = isMacPlatform();
	let altPending = false;
	const onKeyDown = (event: KeyboardEvent) => {
		altPending = event.key === 'Alt' && !event.ctrlKey && !event.metaKey && !event.shiftKey;
		if (event.defaultPrevented || event.isComposing) return;
		const shortcut = findShortcut(event, mac);
		if (!shortcut) return;
		const typing = inTextField(event) || (ctx.grid()?.isEditing() ?? false);
		if (typing && !shortcut.global) return;
		event.preventDefault();
		if (shortcut.action) actions(shortcut.action);
		else if (shortcut.command) void ctx.commands.run(shortcut.command, shortcut.arg);
	};
	const onKeyUp = (event: KeyboardEvent) => {
		const tapped = altPending && event.key === 'Alt';
		altPending = false;
		if (tapped && !mac) {
			event.preventDefault();
			actions('keytips');
		}
	};
	host.addEventListener('keydown', onKeyDown);
	host.addEventListener('keyup', onKeyUp);
	return () => {
		host.removeEventListener('keydown', onKeyDown);
		host.removeEventListener('keyup', onKeyUp);
	};
}
