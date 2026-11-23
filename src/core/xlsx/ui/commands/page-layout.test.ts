import { describe, expect, it } from 'vitest';
import {
	createEditSession,
	createWorkbook,
	THEME_COLOR_PRESETS,
	THEME_FONT_PRESETS,
	THEME_PRESETS,
} from '../../index';
import type { EditorContext } from '../context';
import { themeFontItems, themePaletteItems, themePresetItems } from '../ribbon/tabs/theme-palette';
import { pageLayoutCommands } from './page-layout';

function makeContext(withWorkbook = true) {
	const wb = createWorkbook();
	const session = createEditSession(wb);
	const toasts: string[] = [];
	const ctx = {
		workbook: () => (withWorkbook ? wb : undefined),
		session: () => (withWorkbook ? session : undefined),
		activeSheet: () => 0,
		readOnly: () => false,
		selection: {
			get: () => ({
				sheet: 0,
				active: { row: 0, col: 0 },
				anchor: { row: 0, col: 0 },
				ranges: [{ start: { row: 0, col: 0 }, end: { row: 0, col: 0 } }],
			}),
		},
		t: (text: string) => text,
		toast: (msg: string) => toasts.push(msg),
	} as unknown as EditorContext;
	return { wb, session, ctx, toasts };
}

describe('Page Layout theme commands and galleries', () => {
	it('exposes built-in themes, colours and fonts in gallery items', () => {
		const { ctx } = makeContext();
		expect(themePresetItems(ctx)).toHaveLength(THEME_PRESETS.length);
		expect(themePaletteItems(ctx)).toHaveLength(THEME_COLOR_PRESETS.length);
		expect(themeFontItems(ctx)).toHaveLength(THEME_FONT_PRESETS.length);
		expect(themePresetItems(ctx)[0]?.preview).toContain('<svg');
		expect(themePaletteItems(ctx)[0]?.preview).toContain('<svg');
		expect(themeFontItems(ctx)[0]?.preview).toContain('font-family');
	});

	it('applies theme preset, colour scheme, and font scheme via commands with undo', () => {
		const { wb, session, ctx, toasts } = makeContext();
		const commands = new Map(pageLayoutCommands().map((cmd) => [cmd.id, cmd]));

		const themesCmd = commands.get('page.themes')!;
		const colorsCmd = commands.get('page.theme-colors')!;
		const fontsCmd = commands.get('page.theme-fonts')!;

		expect(themesCmd.enabled?.(ctx)).toBe(true);
		expect(colorsCmd.enabled?.(ctx)).toBe(true);
		expect(fontsCmd.enabled?.(ctx)).toBe(true);

		themesCmd.run(ctx, 'facet');
		expect(wb.theme.colors[4]).toBe('90C226');
		expect(wb.theme.majorFont).toBe('Trebuchet MS');
		expect(toasts).toHaveLength(0);

		colorsCmd.run(ctx, 'ion');
		expect(wb.theme.colors[4]).toBe('B01513');
		expect(wb.theme.majorFont).toBe('Trebuchet MS');

		fontsCmd.run(ctx, 'georgia');
		expect(wb.theme.colors[4]).toBe('B01513');
		expect(wb.theme.majorFont).toBe('Georgia');
		expect(wb.theme.minorFont).toBe('Georgia');

		expect(session.undo()).toBe(true);
		expect(wb.theme.majorFont).toBe('Trebuchet MS');
	});

	it('marks theme commands as editing and disables them without a workbook', () => {
		const { ctx } = makeContext(false);
		const commands = new Map(pageLayoutCommands().map((cmd) => [cmd.id, cmd]));
		for (const id of ['page.themes', 'page.theme-colors', 'page.theme-fonts']) {
			const cmd = commands.get(id)!;
			expect(cmd.editing).toBe(true);
			expect(cmd.enabled?.(ctx)).toBe(false);
		}
	});
});
