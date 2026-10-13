import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { parseVsdx } from './parser';
import { editVsdx } from './edit';

const native = () =>
	readFileSync(new URL('./__fixtures__/inherited-field-native.vsdx', import.meta.url));
const texts = async (bytes: Uint8Array) =>
	(await parseVsdx(bytes)).pages[0]!.shapes.map((shape) => shape.text.plainText);

describe('native inherited text fields', () => {
	it('uses each instance string cache instead of the master placeholder', async () => {
		expect(await texts(native())).toEqual(['42', 'Ready', '73.50', '']);
	});
	it('preserves cached field values through unrelated edits and reopen', async () => {
		const source = native();
		const page = (await parseVsdx(source)).pages[0]!;
		const result = await editVsdx(source, [
			{ type: 'rename-page', pageId: page.id, name: 'Metrics' },
		]);
		expect(await texts(result.bytes)).toEqual(['42', 'Ready', '73.50', '']);
	});
	it('keeps the stored display for unimplemented internal pictures', async () => {
		const zip = await JSZip.loadAsync(native());
		for (const part of Object.values(zip.files)) {
			if (part.name.startsWith('visio/masters/') && part.name.endsWith('.xml')) {
				zip.file(
					part.name,
					(await part.async('string')).replaceAll('FIELDPICTURE(0)', 'FIELDPICTURE(999)'),
				);
			}
		}
		expect(await texts(await zip.generateAsync({ type: 'uint8array' }))).toEqual([
			'Abc',
			'Abc',
			'Abc',
			'Abc',
		]);
	});
});
