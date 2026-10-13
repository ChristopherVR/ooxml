import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { editVsdx } from './edit';
import { parseVsdx } from './parser';

it('writes native DATE cache syntax and string format units for inserted date fields', async () => {
	const bytes = readFileSync(
		new URL('./__fixtures__/date-field-write-source.vsdx', import.meta.url),
	);
	const source = (await parseVsdx(bytes)).pages[0]!;
	const saved = await editVsdx(bytes, [
		{
			type: 'insert-text-field',
			pageId: source.id,
			shapeId: source.shapes[0]!.id,
			formula: 'DOCCREATION()',
			format: '{{yyyy-MM-dd}}',
		},
	]);
	const originalZip = await JSZip.loadAsync(bytes),
		zip = await JSZip.loadAsync(saved.bytes);
	const xml = await zip.file('visio/pages/page1.xml')!.async('string');
	expect(xml).toMatch(
		/N="Value" V="\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}\.\d{3}" U="DATE" F="DOCCREATION\(\)"/,
	);
	expect(xml).toContain('<Cell N="Format" V="{{yyyy-MM-dd}}" U="STR"');
	expect(await zip.file('docProps/core.xml')!.async('string')).toBe(
		await originalZip.file('docProps/core.xml')!.async('string'),
	);
	const reopened = (await parseVsdx(saved.bytes)).pages[0]!.shapes[0]!.text;
	expect(reopened.plainText).toMatch(/^Created: \d{4}-\d{2}-\d{2}$/);
	expect(reopened.fields).toMatchObject([{ formula: 'DOCCREATION()' }]);
});
