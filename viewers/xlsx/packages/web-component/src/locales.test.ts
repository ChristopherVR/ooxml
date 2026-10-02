// @vitest-environment jsdom
// Locale coverage: every locale defines every English key with the same placeholders, and every
// user-facing string of the editor (t() literals, command labels, ribbon tab and group labels,
// shortcut labels) is a key of the English table.
import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join, relative } from 'node:path';
import { describe, expect, it } from 'vitest';
import { EDITOR_LOCALES, STRINGS, normalizeEditorLocale } from './localization';
import { SHORTCUTS } from './keyboard';
import { shellCommands } from './shell-commands';
import { allCommands, commandTabs } from './commands/index';
import type { RibbonControl } from './ribbon/parts';

const SRC = join(process.cwd(), 'packages/web-component/src');
const english = STRINGS.en;
const placeholders = (text: string) => [...text.matchAll(/\{\w+\}/g)].map((m) => m[0]).sort();
/** Terms that are legitimately the same as English in some locale (loanwords, names, symbols). */
const MAY_MATCH =
	/^(OK|Zoom|Normal|Format|Info|Total|Budget|Status|Transport|Description|Date|Minimum|Maximum|Options|Orange|Design|Text \d|Subtotal|General|Manual|Autor|Author)/i;

function sourceFiles(dir: string): string[] {
	return readdirSync(dir).flatMap((name) => {
		const path = join(dir, name);
		if (statSync(path).isDirectory())
			return name === 'locales' || name === 'test-support' ? [] : sourceFiles(path);
		return /\.ts$/.test(name) && !/\.test\.ts$|test-support|test-context/.test(name) ? [path] : [];
	});
}

/** Literal keys passed to a translator: `t('...')`, `ctx.t('...')`, `this.ctx.t(\n'...')`. */
function literalKeys(): Array<{ file: string; key: string }> {
	const found: Array<{ file: string; key: string }> = [];
	for (const file of sourceFiles(SRC)) {
		const text = readFileSync(file, 'utf8');
		for (const match of text.matchAll(/\bt\(\s*'((?:[^'\\]|\\.)*)'/g))
			found.push({ file: relative(SRC, file), key: match[1]!.replace(/\\'/g, "'") });
	}
	return found;
}

function controlLabels(control: RibbonControl): string[] {
	if (control.kind === 'stack') return control.controls.flatMap(controlLabels);
	if (control.kind === 'menu')
		return [
			control.label,
			...control.items.flatMap((item) => ('label' in item && item.label ? [item.label] : [])),
		];
	if (control.kind === 'split')
		return control.menu.flatMap((item) => ('label' in item && item.label ? [item.label] : []));
	if (control.kind === 'color') return [control.label];
	return [];
}

describe('locale catalogue', () => {
	it('lists the five locales', () => {
		expect([...EDITOR_LOCALES]).toEqual(['en', 'fr', 'de', 'es', 'zh-CN']);
		expect(Object.keys(STRINGS).sort()).toEqual([...EDITOR_LOCALES].sort());
	});

	for (const locale of EDITOR_LOCALES)
		describe(locale, () => {
			it('defines every key, without extras or empty strings, keeping placeholders', () => {
				const table = STRINGS[locale];
				expect(
					Object.keys(table).filter((key) => !(key in english)),
					'extra keys',
				).toEqual([]);
				const missing = Object.keys(english).filter(
					(key) => typeof table[key] !== 'string' || !table[key]!.trim(),
				);
				expect(missing, 'missing keys').toEqual([]);
				const broken = Object.keys(english).filter(
					(key) =>
						JSON.stringify(placeholders(table[key] ?? '')) !==
						JSON.stringify(placeholders(english[key]!)),
				);
				expect(broken, 'placeholders').toEqual([]);
			});

			if (locale !== 'en')
				it('does not leave longer UI text in English', () => {
					const same = Object.keys(english).filter(
						(key) =>
							STRINGS[locale][key] === english[key] &&
							key.split(' ').length > 2 &&
							!MAY_MATCH.test(key) &&
							/^[a-z]/i.test(key),
					);
					expect(same).toEqual([]);
				});
		});

	it('has English identity text and no em-dashes', () => {
		for (const [key, value] of Object.entries(english))
			if (!key.includes('.')) expect(value, key).toBe(key);
		for (const locale of EDITOR_LOCALES)
			for (const value of Object.values(STRINGS[locale]))
				expect(value).not.toContain(String.fromCharCode(0x2014));
	});
});

describe('coverage of the editor strings', () => {
	it('every literal t() key is in the English table', () => {
		const missing = literalKeys().filter(({ key }) => !(key in english) && !key.includes('${'));
		expect(missing.map(({ file, key }) => `${file}: ${key}`)).toEqual([]);
	});

	it('every command label, ribbon tab, group and menu label, and shortcut label is translated', () => {
		const labels = new Set<string>();
		for (const command of [
			...allCommands(),
			...shellCommands({
				fileCommand() {},
				openBackstage() {},
				showShortcuts() {},
				formulaBarShown: () => true,
				setFormulaBarShown() {},
			}),
		])
			labels.add(command.label);
		for (const tab of commandTabs()) {
			labels.add(tab.label);
			for (const group of tab.groups) {
				labels.add(group.label);
				for (const control of group.controls)
					for (const label of controlLabels(control)) labels.add(label);
			}
		}
		for (const shortcut of SHORTCUTS) labels.add(shortcut.label);
		// Function names (SUM, XLOOKUP) are not translated, as in Excel's English function names.
		const missing = [...labels].filter(
			(label) => !(label in english) && !/^[A-Z][A-Z0-9.]*$/.test(label),
		);
		expect(missing.sort()).toEqual([]);
	});

	it('normalizes BCP 47 tags', () => {
		for (const [input, expected] of [
			['fr-CA', 'fr'],
			['de_AT', 'de'],
			['es-MX', 'es'],
			['zh-Hans', 'zh-CN'],
			['zh-TW', 'en'],
			['ja', 'en'],
			[undefined, 'en'],
		] as const)
			expect(normalizeEditorLocale(input)).toBe(expected);
	});
});
