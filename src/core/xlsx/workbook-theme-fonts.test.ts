/**
 * A new workbook takes its theme fonts from the one shared default (`OFFICE_THEME_FONTS`) or
 * the caller's, and its Normal style names the theme body font with `scheme="minor"`, as Excel's
 * own new workbooks do. Two saves of the same workbook are byte-identical (1980 ZIP dates).
 */
import { afterEach, describe, expect, it, vi } from 'vitest';
import { OFFICE_THEME_FONTS } from '../drawingml/office-theme-fonts';
import { createWorkbook } from './workbook';
import { saveXlsx } from './write/index';

afterEach(() => {
	vi.useRealTimers();
});

describe('new workbook theme fonts', () => {
	it('defaults to the shared theme fonts', () => {
		const book = createWorkbook();
		expect(book.theme.majorFont).toBe(OFFICE_THEME_FONTS.major);
		expect(book.theme.minorFont).toBe(OFFICE_THEME_FONTS.minor);
		expect(book.styles[0]?.font).toMatchObject({ name: OFFICE_THEME_FONTS.minor, scheme: 'minor' });
	});

	it('takes the caller theme fonts', () => {
		const book = createWorkbook({ themeFonts: { major: 'Head', minor: 'Body' } });
		expect(book.theme.majorFont).toBe('Head');
		expect(book.theme.minorFont).toBe('Body');
		expect(book.styles[0]?.font?.name).toBe('Body');
		expect(book.namedStyles?.[0]?.style.font?.name).toBe('Body');
	});

	it('saves the same bytes at different times', async () => {
		vi.useFakeTimers({ toFake: ['Date'] });
		vi.setSystemTime(new Date('2026-05-05T10:00:00Z'));
		const first = await saveXlsx(createWorkbook());
		vi.setSystemTime(new Date('2027-01-02T03:04:05Z'));
		const second = await saveXlsx(createWorkbook());
		expect(Buffer.from(second).equals(Buffer.from(first))).toBe(true);
	});
});
