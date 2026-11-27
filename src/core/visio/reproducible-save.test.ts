/**
 * A Visio edit with an injected `now` gives the same bytes whenever it runs: the recordset's
 * `TimeRefreshed` is that time and every ZIP entry is dated 1980-01-01, as Visio writes. A new
 * drawing names the shared default theme body font, or the one the caller gives.
 */
import JSZip from 'jszip';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { OFFICE_THEME_FONTS } from '../drawingml/office-theme-fonts';
import { createVsdx } from './create-document';
import type { VisioDataColumn } from './data-recordsets';
import { editVsdx } from './edit';

const NOW = new Date('2026-01-15T09:00:00Z');
const columns: VisioDataColumn[] = [{ name: 'Name', label: 'Name', type: 'string' }];

async function build(): Promise<Uint8Array> {
	const drawing = await editVsdx(await createVsdx(), [
		{ type: 'create-rectangle', pageId: '0', shapeId: '1', x: 2, y: 2, width: 2, height: 1 },
	]);
	return (
		await editVsdx(
			drawing.bytes,
			[{ type: 'import-data-recordset', pageId: '0', name: 'a.csv', columns, rows: [['Web']] }],
			{ now: NOW },
		)
	).bytes;
}

afterEach(() => {
	vi.useRealTimers();
});

describe('reproducible Visio edits', () => {
	it('gives identical bytes for the same now at different times', async () => {
		vi.useFakeTimers({ toFake: ['Date'] });
		vi.setSystemTime(new Date('2026-05-05T10:00:00Z'));
		const first = await build();
		vi.setSystemTime(new Date('2026-06-07T11:00:03Z'));
		const second = await build();
		expect(Buffer.from(second).equals(Buffer.from(first))).toBe(true);
		const zip = await JSZip.loadAsync(first);
		const recordsets = await zip.file('visio/data/recordsets.xml')!.async('string');
		expect(recordsets).toContain('TimeRefreshed="2026-01-15T09:00:00"');
	});

	it('dates every entry 1980-01-01', async () => {
		const bytes = await createVsdx();
		const view = new DataView(bytes.buffer, bytes.byteOffset);
		expect(view.getUint32(0, true)).toBe(0x04034b50);
		expect(view.getUint16(10, true)).toBe(0);
		expect(view.getUint16(12, true)).toBe(0x21);
	});

	it('names the default theme body font, or the one given', async () => {
		const read = async (bytes: Uint8Array) =>
			(await JSZip.loadAsync(bytes)).file('visio/document.xml')!.async('string');
		expect(await read(await createVsdx())).toContain(`Name="${OFFICE_THEME_FONTS.minor}"`);
		const custom = await read(await createVsdx({ themeFonts: { minor: 'Body Face' } }));
		expect(custom).toContain('Name="Body Face"');
		expect(custom).not.toContain(OFFICE_THEME_FONTS.minor);
	});
});
