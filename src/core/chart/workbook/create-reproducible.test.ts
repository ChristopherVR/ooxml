/**
 * A new chart workbook is the same bytes whenever it is written (its ZIP entries are dated
 * 1980-01-01, as Office writes) and names the host theme's body font on its Normal style.
 */
import JSZip from 'jszip';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { OFFICE_THEME_FONTS } from '../../drawingml/office-theme-fonts';
import { createChartWorkbookFromRows } from './create';

const rows = [
	[null, 'Sales'],
	['Q1', 4],
	['Q2', 7],
];

afterEach(() => {
	vi.useRealTimers();
});

describe('createChartWorkbookFromRows', () => {
	it('is byte-identical across runs at different times', async () => {
		vi.useFakeTimers({ toFake: ['Date'] });
		vi.setSystemTime(new Date('2026-05-05T10:00:00Z'));
		const first = await createChartWorkbookFromRows(rows);
		vi.setSystemTime(new Date('2027-01-02T03:04:05Z'));
		const second = await createChartWorkbookFromRows(rows);
		expect(Buffer.from(second).equals(Buffer.from(first))).toBe(true);
		const view = new DataView(first.buffer, first.byteOffset);
		expect(view.getUint16(12, true)).toBe(0x21);
	});

	it('names the default theme body font, or the host theme one, as the minor scheme font', async () => {
		const styles = async (bytes: Uint8Array) =>
			(await JSZip.loadAsync(bytes)).file('xl/styles.xml')!.async('string');
		expect(await styles(await createChartWorkbookFromRows(rows))).toContain(
			`<name val="${OFFICE_THEME_FONTS.minor}"/><family val="2"/><scheme val="minor"/>`,
		);
		const custom = await styles(
			await createChartWorkbookFromRows(rows, 'Sheet1', { minorFont: 'Body & Face' }),
		);
		expect(custom).toContain('<name val="Body &amp; Face"/>');
	});
});
