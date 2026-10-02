// @vitest-environment jsdom
import { createWorkbook } from '@christophervr/xlsx-core';
import { afterEach, describe, expect, it } from 'vitest';
import { NAVIGATION_STRINGS } from '../commands/i18n/navigation.js';
import { SHORTCUTS } from '../keyboard.js';
import { clickButton, createTestContext, dialogEl } from '../commands/test-support.js';
import { registerNavigationDialogs } from './register-navigation.js';

afterEach(() => (document.body.innerHTML = ''));

const DIALOGS: ReadonlyArray<readonly [string, unknown?]> = [
	['find-replace', { tab: 'replace' }],
	['go-to'],
	['go-to-special'],
	['insert-function', { category: 'all' }],
	['name-manager'],
	['define-name'],
	['create-names'],
	['workbook-statistics'],
	['comments-list'],
	['shortcut-help'],
	['feature-status'],
];

describe('navigation dialog strings', () => {
	it('translates every string the dialogs show, in all four locales', async () => {
		const ctx = createTestContext(createWorkbook());
		const used = new Set<string>();
		const t = ctx.t;
		ctx.t = (key, vars) => {
			used.add(key);
			return t(key, vars);
		};
		registerNavigationDialogs(ctx);
		ctx.session()!.setDefinedName({ name: 'N', formula: '1' });
		for (const [name, props] of DIALOGS) {
			const done = ctx.dialogs.open(name, props);
			const dialog = dialogEl(ctx, name);
			for (const b of dialog.querySelectorAll<HTMLButtonElement>('button'))
				if (/Options/.test(b.textContent ?? '')) b.click();
			clickButton(
				dialog,
				[...dialog.querySelectorAll('button')].some((b) => b.textContent === 'Cancel')
					? 'Cancel'
					: 'Close',
			);
			await done;
		}
		// Shortcut labels come from the shell's keyboard map and are translated by the shell.
		const shell = new Set(SHORTCUTS.map((s) => s.label));
		const missing = [...used].filter(
			(key) => !NAVIGATION_STRINGS[key] && !shell.has(key) && !/^[A-Z][A-Z0-9.]+$/.test(key),
		);
		expect(missing).toEqual([]);
		for (const [key, texts] of Object.entries(NAVIGATION_STRINGS)) {
			expect(texts.length, key).toBe(4);
			for (const text of texts) expect(text, key).not.toContain(String.fromCharCode(0x2014));
		}
	});
});
