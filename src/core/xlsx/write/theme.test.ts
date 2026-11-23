import { describe, expect, it } from 'vitest';
import { createWorkbook } from '../workbook';
import { createEditSession } from '../edit/session';
import { loadXlsx } from '../read/load';
import { findThemeColorPreset, findThemeFontPreset, findThemePreset } from '../theme-presets';
import { saveXlsx } from './save';
import { defaultThemeXml, themePart } from './theme';

describe('theme XML writing and round-trip', () => {
	it('round-trips a changed theme preset through save and reload', async () => {
		const wb = createWorkbook();
		const session = createEditSession(wb);
		const integral = findThemePreset('integral')!;
		session.setTheme({
			colors: integral.colors,
			majorFont: integral.majorFont,
			minorFont: integral.minorFont,
		});

		const bytes = await saveXlsx(wb);
		const reloaded = await loadXlsx(bytes);

		expect(reloaded.theme.colors).toEqual([...integral.colors]);
		expect(reloaded.theme.majorFont).toBe(integral.majorFont);
		expect(reloaded.theme.minorFont).toBe(integral.minorFont);
	});

	it('patches an existing theme part on save without losing extensions', async () => {
		const wb = createWorkbook();
		const initialBytes = await saveXlsx(wb);
		const loaded = await loadXlsx(initialBytes);
		const session = createEditSession(loaded);

		const colors = findThemeColorPreset('organic')!;
		const fonts = findThemeFontPreset('georgia')!;
		session.setTheme({
			colors: colors.colors,
			majorFont: fonts.majorFont,
			minorFont: fonts.minorFont,
		});

		const savedBytes = await saveXlsx(loaded);
		const reloaded = await loadXlsx(savedBytes);

		expect(reloaded.theme.colors).toEqual([...colors.colors]);
		expect(reloaded.theme.majorFont).toBe('Georgia');
		expect(reloaded.theme.minorFont).toBe('Georgia');
	});

	it('returns undefined when the theme matches sourceXml and preserves clrScheme extLst when patching', () => {
		const wb = createWorkbook();
		const source = defaultThemeXml(wb.theme).replace(
			'</a:clrScheme>',
			'<a:extLst><a:ext uri="{test}"/></a:extLst></a:clrScheme>',
		);
		expect(themePart(wb.theme, source)).toBeUndefined();

		const patched = themePart(
			{
				...wb.theme,
				colors: [...findThemeColorPreset('ion')!.colors],
				majorFont: 'Verdana',
				minorFont: 'Verdana',
			},
			source,
		);
		expect(patched).toBeDefined();
		expect(patched).toContain('<a:extLst><a:ext uri="{test}"/></a:extLst>');
		expect(patched).toContain('val="B01513"');
		expect(patched).toContain('typeface="Verdana"');
	});
});
