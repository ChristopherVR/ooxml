import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { createEditSession } from '../edit/index.js';
import { loadXlsx } from '../read/index.js';
import { createWorkbook } from '../workbook.js';
import { saveXlsx } from './index.js';

const fixture = (name: string) =>
	new Uint8Array(readFileSync(path.join(import.meta.dirname, '..', '__fixtures__', name)));

const text = async (bytes: Uint8Array, part: string) =>
	(await JSZip.loadAsync(bytes)).file(part)?.async('string');

describe('document properties round trip', () => {
	it('keeps an Excel app.xml unchanged when nothing was edited', async () => {
		const bytes = fixture('excel-features.xlsx');
		const before = await text(bytes, 'docProps/app.xml');
		const saved = await saveXlsx(await loadXlsx(bytes));
		const after = await text(saved, 'docProps/app.xml');
		// The serializer writes an empty element as `<Company/>`.
		const strip = (xml: string | undefined) =>
			(xml ?? '').replace(/<\?xml[^>]*\?>\s*/, '').replace(/<(\w+)><\/\1>/g, '<$1/>');
		expect(strip(after)).toBe(strip(before));
	});

	it('writes core, app and custom properties set through the edit session, with undo', async () => {
		const wb = await loadXlsx(fixture('excel-features.xlsx'));
		const session = createEditSession(wb);
		expect(
			session.setDocumentProperties({
				title: 'Quarterly',
				category: 'Finance',
				contentStatus: 'Draft',
				language: 'en-ZA',
				manager: 'Mia Manager',
				hyperlinkBase: 'https://example.com/',
				custom: [
					{ name: 'Project', type: 'lpwstr', value: 'Apollo' },
					{ name: 'Budget', type: 'r8', value: 1250.5 },
					{ name: 'Approved', type: 'bool', value: true },
					{ name: 'Count', type: 'i4', value: 7 },
					{ name: 'Due', type: 'filetime', value: '2026-10-03T00:00:00Z' },
				],
			}),
		).toBe(true);
		const saved = await saveXlsx(wb);
		const types = (await text(saved, '[Content_Types].xml')) ?? '';
		expect(types).toContain(
			'PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"',
		);
		expect(await text(saved, '_rels/.rels')).toContain('relationships/custom-properties');
		const back = await loadXlsx(saved);
		expect(back.properties).toMatchObject({
			title: 'Quarterly',
			category: 'Finance',
			contentStatus: 'Draft',
			language: 'en-ZA',
			manager: 'Mia Manager',
			hyperlinkBase: 'https://example.com/',
		});
		expect(back.properties.custom?.map((p) => [p.name, p.pid])).toEqual([
			['Project', 2],
			['Budget', 3],
			['Approved', 4],
			['Count', 5],
			['Due', 6],
		]);

		session.undo();
		expect(wb.properties.title).not.toBe('Quarterly');
		expect(wb.properties.custom).toBeUndefined();
		session.redo();
		expect(wb.properties.manager).toBe('Mia Manager');
		expect(session.setDocumentProperties({ manager: 'Mia Manager' })).toBe(false);
	});

	it('clears fields with null and removes custom.xml for an empty list', async () => {
		const wb = createWorkbook();
		const session = createEditSession(wb);
		session.setDocumentProperties({
			title: 'T',
			custom: [{ name: 'X', type: 'lpwstr', value: 'y' }],
		});
		const first = await loadXlsx(await saveXlsx(wb));
		expect(first.properties.custom).toHaveLength(1);
		const again = createEditSession(first);
		again.setDocumentProperties({ title: null, custom: [] });
		const saved = await saveXlsx(first);
		expect(await text(saved, 'docProps/custom.xml')).toBeUndefined();
		expect(await text(saved, '_rels/.rels')).not.toContain('custom-properties');
		const back = await loadXlsx(saved);
		expect(back.properties.title).toBeUndefined();
		expect(back.properties.custom).toBeUndefined();
	});

	it('carries a source custom.xml the model did not touch and refreshes sheet titles', async () => {
		const wb = createWorkbook();
		createEditSession(wb).setDocumentProperties({
			custom: [{ name: 'Keep', type: 'i4', value: 1 }],
		});
		const loaded = await loadXlsx(await saveXlsx(wb));
		delete loaded.properties.custom;
		const first = loaded.sheets[0];
		if (first) first.name = 'Renamed';
		const saved = await saveXlsx(loaded);
		expect(await text(saved, 'docProps/custom.xml')).toContain('name="Keep"');
		const back = await loadXlsx(saved);
		expect(back.properties.titlesOfParts).toEqual(['Renamed']);
		expect(back.properties.headingPairs).toEqual([{ name: 'Worksheets', count: 1 }]);
	});
});
