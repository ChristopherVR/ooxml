import { expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { loadXlsx } from './index';
import { saveXlsx } from '../write/index';
import { createEditSession } from '../edit/session';
import { setDrawingAnchor } from '../edit/drawings';
import { testContext } from '../edit/test-context';
it('reads and edits picture crop with history and preserves opaque picture content on move/save', async () => {
	const zip = await JSZip.loadAsync(
		readFileSync(new URL('../__fixtures__/excel-features.xlsx', import.meta.url)),
	);
	const part = Object.keys(zip.files).find(
		(name) => /^xl\/drawings\/drawing\d+\.xml$/.test(name) && zip.file(name),
	)!;
	const source = await zip.file(part)!.async('string');
	expect(source).toContain('<xdr:pic>');
	zip.file(
		part,
		source
			.replace('<a:blip ', '<a:blip xmlns:v="urn:vendor" v:opaque="keep" ')
			.replace(
				'</a:blip>',
				'<a:extLst xmlns:v="urn:vendor"><a:ext uri="keep"><v:opaque/></a:ext></a:extLst></a:blip>',
			)
			.replace('<a:stretch>', '<a:srcRect l="25000" r="-10000"/><a:stretch>'),
	);
	const wb = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
	const sheetIndex = wb.sheets.findIndex((sheet) => sheet.drawings.some((d) => d.kind === 'image'));
	const index = wb.sheets[sheetIndex]!.drawings.findIndex((d) => d.kind === 'image');
	const image = wb.sheets[sheetIndex]!.drawings[index]!;
	expect(image.kind).toBe('image');
	if (image.kind !== 'image') return;
	expect(image.crop).toEqual({ cropLeft: 0.25, cropRight: -0.1 });
	const session = createEditSession(wb);
	session.setPictureCrop(sheetIndex, index, { cropLeft: 0.1, cropRight: 0.2 });
	expect(image.crop).toEqual({ cropLeft: 0.1, cropRight: 0.2 });
	session.undo();
	const restored = wb.sheets[sheetIndex]!.drawings[index]!;
	expect(restored.kind === 'image' && restored.crop).toEqual({ cropLeft: 0.25, cropRight: -0.1 });
	session.setPictureCrop(sheetIndex, index, { cropLeft: 0.1, cropRight: 0.2 });
	const ctx = testContext(wb);
	setDrawingAnchor(ctx, sheetIndex, index, {
		from: { row: 15, col: 2, rowOffset: 0, colOffset: 0 },
		ext: { cx: 1905000, cy: 952500 },
	});
	const saved = await saveXlsx(wb);
	const xml = await (await JSZip.loadAsync(saved)).file(part)!.async('string');
	expect(xml).toContain('l="10000"');
	expect(xml).toContain('r="20000"');
	expect(xml).toContain('v:opaque="keep"');
	const again = await loadXlsx(saved);
	expect(again.sheets[sheetIndex]!.drawings[index]).toMatchObject({
		kind: 'image',
		crop: { cropLeft: 0.1, cropRight: 0.2 },
		anchor: { from: { row: 15 } },
	});
	expect(() => session.setPictureCrop(sheetIndex, index, { cropLeft: 1 })).toThrow(RangeError);
});
