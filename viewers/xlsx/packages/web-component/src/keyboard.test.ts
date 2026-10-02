// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import {
	findShortcut,
	formatKeys,
	installKeyboard,
	matchesKeys,
	SHORTCUTS,
	type ShellAction,
} from './keyboard';
import { shortcutRows } from './shortcut-help';
import { shellFixture, spyCommand } from './test-support/shell';

const key = (key: string, mods: Partial<KeyboardEventInit> = {}) => ({
	key,
	code: '',
	ctrlKey: false,
	metaKey: false,
	altKey: false,
	shiftKey: false,
	...mods,
});

describe('key specs', () => {
	it('matches modifiers exactly, with Mod as Ctrl or Cmd', () => {
		expect(matchesKeys(key('s', { ctrlKey: true }), 'Mod+S', false)).toBe(true);
		expect(matchesKeys(key('s', { metaKey: true }), 'Mod+S', true)).toBe(true);
		expect(matchesKeys(key('s', { ctrlKey: true }), 'Mod+S', true)).toBe(false);
		expect(matchesKeys(key('s', { ctrlKey: true, shiftKey: true }), 'Mod+S', false)).toBe(false);
		expect(matchesKeys(key('+', { ctrlKey: true }), 'Mod++', false)).toBe(true);
		expect(matchesKeys(key(' ', { shiftKey: true }), 'Shift+ ', false)).toBe(true);
		expect(
			matchesKeys(
				{ ...key('4', { ctrlKey: true, shiftKey: true }), code: 'Digit4' },
				'Mod+Shift+code:Digit4',
				false,
			),
		).toBe(true);
	});

	it('maps Excel shortcuts to command ids', () => {
		const id = (event: ReturnType<typeof key>) =>
			findShortcut(event, false)?.command ?? findShortcut(event, false)?.action;
		expect(id(key('b', { ctrlKey: true }))).toBe('home.bold');
		expect(id(key('1', { ctrlKey: true }))).toBe('format.cells');
		expect(id(key('$', { ctrlKey: true, shiftKey: true }))).toBe('home.format-currency');
		expect(id(key('=', { altKey: true }))).toBe('home.autosum');
		expect(id(key('PageDown', { ctrlKey: true }))).toBe('sheet.next');
		expect(id(key('L', { ctrlKey: true, shiftKey: true }))).toBe('data.filter');
		expect(id(key('F6'))).toBe('next-region');
		expect(id(key('/', { ctrlKey: true }))).toBe('help');
		expect(id(key('F5'))).toBe('home.go-to');
	});

	it('formats keys for help', () => {
		expect(formatKeys('Mod+Shift+L', false)).toBe('Ctrl+Shift+L');
		expect(formatKeys('Mod+S', true)).toBe('Cmd+S');
		expect(formatKeys('Shift+ ', false)).toBe('Shift+Space');
		expect(formatKeys('Mod++', false)).toBe('Ctrl++');
	});

	it('has no duplicate key specs', () => {
		const specs = SHORTCUTS.flatMap((shortcut) => shortcut.keys);
		expect(new Set(specs).size).toBe(specs.length);
	});
});

describe('installKeyboard', () => {
	function setup() {
		const fixture = shellFixture();
		const bold = spyCommand('home.bold');
		const save = spyCommand('file.save');
		fixture.core.commands.registerAll([bold, save]);
		const actions: ShellAction[] = [];
		installKeyboard(fixture.core.ctx, (action) => actions.push(action));
		const input = document.createElement('input');
		fixture.core.ctx.root.append(input);
		return { ...fixture, bold, save, actions, input };
	}
	const press = (target: EventTarget, init: KeyboardEventInit) => {
		const event = new KeyboardEvent('keydown', {
			bubbles: true,
			composed: true,
			cancelable: true,
			...init,
		});
		target.dispatchEvent(event);
		return event;
	};

	it('runs commands and shell actions, preventing the browser default', () => {
		const { host, bold, actions } = setup();
		expect(press(host, { key: 'b', ctrlKey: true }).defaultPrevented).toBe(true);
		expect(bold.runs).toHaveLength(1);
		press(host, { key: 'F1' });
		expect(actions).toEqual(['help']);
	});

	it('leaves handled events and typing in text fields alone, except global shortcuts', () => {
		const { host, bold, save, input } = setup();
		const handled = new KeyboardEvent('keydown', {
			key: 'b',
			ctrlKey: true,
			bubbles: true,
			cancelable: true,
		});
		handled.preventDefault();
		host.dispatchEvent(handled);
		expect(press(input, { key: 'b', ctrlKey: true }).defaultPrevented).toBe(false);
		press(input, { key: 's', ctrlKey: true });
		expect(bold.runs).toHaveLength(0);
		expect(save.runs).toHaveLength(1);
	});

	it('shows KeyTips on a lone Alt tap', () => {
		const { host, actions } = setup();
		press(host, { key: 'Alt', altKey: true });
		host.dispatchEvent(new KeyboardEvent('keyup', { key: 'Alt', bubbles: true }));
		expect(actions).toEqual(['keytips']);
	});

	it('lists shortcut rows for registered commands, using their labels', () => {
		const { core } = setup();
		const rows = shortcutRows(core.ctx, false);
		expect(rows.find((row) => row.label === 'home.bold')?.keys).toBe('Ctrl+B / Ctrl+2');
		expect(rows.some((row) => row.label === 'Keyboard shortcuts')).toBe(true);
		expect(rows.some((row) => row.label === 'Insert link')).toBe(false);
	});
});
