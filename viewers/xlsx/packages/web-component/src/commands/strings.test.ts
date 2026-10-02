import { describe, expect, it } from 'vitest';
import type { RibbonControl, RibbonMenuItem } from '../ribbon/parts.js';
import { COMMAND_TABLES, commandStrings } from './i18n/index.js';
import { allCommands, commandTabs } from './index.js';
import { BUILTIN_CELL_STYLES } from '@christophervr/xlsx-core';

const known = new Set(COMMAND_TABLES.flatMap((table) => Object.keys(table)));

/** Every translatable key the ribbon shows (function and font names are proper names). */
function ribbonKeys(): string[] {
	const keys = new Set<string>();
	for (const command of allCommands()) keys.add(command.label);
	const menu = (items: RibbonMenuItem[]) => {
		for (const item of items)
			if (!('separator' in item) && item.label && item.command !== 'formulas.function')
				keys.add(item.label);
	};
	const walk = (controls: RibbonControl[]) => {
		for (const control of controls) {
			if (control.kind === 'menu') {
				keys.add(control.label);
				menu(control.items);
			} else if (control.kind === 'split') menu(control.menu);
			else if (control.kind === 'color') keys.add(control.label);
			else if (control.kind === 'stack') walk(control.controls);
		}
	};
	for (const tab of commandTabs()) {
		keys.add(tab.label);
		for (const group of tab.groups) {
			keys.add(group.label);
			walk(group.controls);
		}
	}
	for (const style of BUILTIN_CELL_STYLES) keys.add(style.name);
	return [...keys];
}

describe('command strings', () => {
	it('translates every ribbon label, command and menu item', () => {
		expect(ribbonKeys().filter((key) => !known.has(key))).toEqual([]);
	});

	it('has four non-empty translations per key, keeping placeholders and no em-dash', () => {
		const problems: string[] = [];
		for (const table of COMMAND_TABLES)
			for (const [key, texts] of Object.entries(table)) {
				const holes = (key.match(/\{\w+\}/g) ?? []).sort().join();
				texts.forEach((text, i) => {
					if (!text.trim()) problems.push(`${key} [${i}] empty`);
					if (text.includes(String.fromCharCode(0x2014))) problems.push(`${key} [${i}] em-dash`);
					if ((text.match(/\{\w+\}/g) ?? []).sort().join() !== holes)
						problems.push(`${key} [${i}] placeholders`);
				});
				if (key.includes(String.fromCharCode(0x2014))) problems.push(`${key} em-dash`);
			}
		expect(problems).toEqual([]);
	});

	it('flattens English as the identity and other locales as translations', () => {
		expect(commandStrings('en').Bold).toBe('Bold');
		expect(commandStrings('fr').Bold).toBe('Gras');
		expect(commandStrings('de')['Merge & Center']).toBe('Verbinden und zentrieren');
		expect(commandStrings('zh-CN').Paste).toBe('粘贴');
		expect(Object.keys(commandStrings('es'))).toEqual(Object.keys(commandStrings('en')));
	});
});
