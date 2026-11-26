import { describe, expect, it } from 'vitest';
import type { RibbonControl, RibbonMenuItem } from 'ooxml-core/xlsx/ui';
import { STRINGS, loadEditorStrings } from 'ooxml-core/xlsx/ui';
import { allCommands, commandTabs } from './index';
import { BUILTIN_CELL_STYLES } from 'ooxml-core/xlsx';

const known = new Set(Object.keys(STRINGS.en));

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
			} else if (control.kind === 'split') {
				if (control.label) keys.add(control.label);
				menu(control.menu);
			} else if (control.kind === 'color') keys.add(control.label);
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

	it('flattens English as the identity and other locales as translations', async () => {
		expect(STRINGS.en.Bold).toBe('Bold');
		expect((await loadEditorStrings('fr')).Bold).toBe('Gras');
		expect((await loadEditorStrings('de'))['Merge & Center']).toBe('Verbinden und zentrieren');
		expect((await loadEditorStrings('zh-CN')).Paste).toBe('粘贴');
	});

	it.each([
		['fr', 'Actualiser', 'Actualiser tout'],
		['de', 'Aktualisieren', 'Alle aktualisieren'],
		['es', 'Actualizar', 'Actualizar todo'],
		['zh-CN', '刷新', '全部刷新'],
	] as const)('localizes the PivotTable Refresh actions in %s', async (locale, refresh, all) => {
		const strings = await loadEditorStrings(locale);
		expect(strings.Refresh).toBe(refresh);
		expect(strings['Refresh All']).toBe(all);
	});
});
