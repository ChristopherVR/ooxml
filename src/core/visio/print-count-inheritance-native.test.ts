import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import JSZip from 'jszip';
import { expect, it } from 'vitest';
import evidence from './__fixtures__/print-count-inheritance-native.json';
import { parseVsdx } from './parser';
import { editVsdx } from './edit';
import { visioPrintTileCount } from './print-count';

const sources = await JSZip.loadAsync(
	readFileSync(new URL('./__fixtures__/print-count-inheritance-native.zip', import.meta.url)),
);
it.each(evidence.cases)('native print-default precedence: $id', async (test) => {
	const original = await sources.file(test.file)!.async('uint8array');
	expect(createHash('sha256').update(original).digest('hex')).toBe(test.sha256);
	const page = (await parseVsdx(original)).pages[0]!;
	expect(page.pageSetup?.printCountSettings?.paperKind).toBe(test.id === 'page-paper' ? 9 : 1);
	expect(page.pageSetup?.printCountSettings?.printZoom).toBe(1);
	expect(page.pageSetup?.printFitToPages).toBe(false);
	expect(visioPrintTileCount(page)).toMatchObject({ kind: 'supported', total: test.count });
	expect((await editVsdx(original, [])).bytes).toEqual(original);
	const renamed = await editVsdx(original, [
		{ type: 'rename-page', pageId: page.id, name: 'Precedence preserved' },
	]);
	expect(renamed.changedParts).toEqual(['visio/pages/pages.xml', 'docProps/app.xml']);
	const oldZip = await JSZip.loadAsync(original),
		newZip = await JSZip.loadAsync(renamed.bytes);
	for (const [name, entry] of Object.entries(oldZip.files))
		if (!entry.dir && !['visio/pages/pages.xml', 'docProps/app.xml'].includes(name))
			expect(await newZip.file(name)!.async('uint8array'), name).toEqual(
				await entry.async('uint8array'),
			);
	expect(visioPrintTileCount((await parseVsdx(renamed.bytes)).pages[0]!)).toEqual(
		visioPrintTileCount(page),
	);
});
