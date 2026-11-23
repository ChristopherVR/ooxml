import { describe, expect, it } from 'vitest';
import { createWorkbook } from '../workbook';
import { resolveColor } from '../layout/colors';
import { fontView } from '../layout/style-view';
import { findThemeColorPreset, findThemeFontPreset, findThemePreset } from '../theme-presets';
import { createEditSession } from './session';
import type { WorkbookChange } from './types';

describe('setTheme', () => {
	it('updates theme colours, fonts, and full presets with undo/redo', () => {
		const wb = createWorkbook();
		const session = createEditSession(wb);
		const changes: WorkbookChange[] = [];
		session.onChange((change) => changes.push(change));

		const originalColors = [...wb.theme.colors];
		const originalMajor = wb.theme.majorFont;
		const originalMinor = wb.theme.minorFont;

		const facet = findThemePreset('facet')!;
		expect(
			session.setTheme(
				{ colors: facet.colors, majorFont: facet.majorFont, minorFont: facet.minorFont },
				'Theme: Facet',
			),
		).toBe(true);

		expect(wb.theme.colors).toEqual([...facet.colors]);
		expect(wb.theme.majorFont).toBe('Trebuchet MS');
		expect(wb.theme.minorFont).toBe('Trebuchet MS');
		expect(changes.at(-1)).toMatchObject({ kind: 'view', label: 'Theme: Facet' });
		expect(resolveColor({ theme: 4 }, wb.theme)).toBe('#90C226');
		expect(fontView({ scheme: 'minor' }, wb.theme).family).toBe('Trebuchet MS');

		// Applying the exact same theme is a no-op and does not add a history step.
		const countBeforeNoop = changes.length;
		expect(
			session.setTheme({
				colors: facet.colors,
				majorFont: facet.majorFont,
				minorFont: facet.minorFont,
			}),
		).toBe(false);
		expect(changes.length).toBe(countBeforeNoop);

		// Undo restores the original theme palette and fonts.
		expect(session.undo()).toBe(true);
		expect(wb.theme.colors).toEqual(originalColors);
		expect(wb.theme.majorFont).toBe(originalMajor);
		expect(wb.theme.minorFont).toBe(originalMinor);
		expect(resolveColor({ theme: 4 }, wb.theme)).toBe('#4472C4');
		expect(fontView({ scheme: 'minor' }, wb.theme).family).toBe('Calibri');

		// Redo reapplies the Facet theme.
		expect(session.redo()).toBe(true);
		expect(wb.theme.colors).toEqual([...facet.colors]);
		expect(wb.theme.majorFont).toBe('Trebuchet MS');
		expect(wb.theme.minorFont).toBe('Trebuchet MS');
	});

	it('updates colours only or fonts only without clobbering the other', () => {
		const wb = createWorkbook();
		const session = createEditSession(wb);

		const ionColors = findThemeColorPreset('ion')!;
		expect(session.setTheme({ colors: ionColors.colors })).toBe(true);
		expect(wb.theme.colors[4]).toBe('B01513');
		expect(wb.theme.majorFont).toBe('Calibri Light');
		expect(wb.theme.minorFont).toBe('Calibri');

		const georgia = findThemeFontPreset('georgia')!;
		expect(
			session.setTheme({ majorFont: georgia.majorFont, minorFont: georgia.minorFont }),
		).toBe(true);
		expect(wb.theme.colors[4]).toBe('B01513');
		expect(wb.theme.majorFont).toBe('Georgia');
		expect(wb.theme.minorFont).toBe('Georgia');
	});

	it('normalizes lowercase and #RRGGBB hex colours and ignores invalid entries', () => {
		const wb = createWorkbook();
		const session = createEditSession(wb);
		const colors = wb.theme.colors.map((c, i) => (i === 4 ? '#abcdef' : i === 5 ? 'bad' : c));
		expect(session.setTheme({ colors })).toBe(true);
		expect(wb.theme.colors[4]).toBe('ABCDEF');
		expect(wb.theme.colors[5]).toBe('ED7D31');
	});
});
