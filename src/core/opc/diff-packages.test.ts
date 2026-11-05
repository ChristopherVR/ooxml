import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { diffOpcPackages } from './diff-packages';

async function pack(parts: Record<string, string>, compression: 'STORE' | 'DEFLATE' = 'DEFLATE') {
	const zip = new JSZip();
	for (const [name, text] of Object.entries(parts)) zip.file(name, text);
	return zip.generateAsync({ type: 'uint8array', compression });
}

describe('diffOpcPackages', () => {
	it('lists added, removed and changed parts', async () => {
		const base = await pack({ 'a.xml': '<a/>', 'b.xml': '<b/>', 'c.xml': '<c/>' });
		const other = await pack({ 'a.xml': '<a/>', 'b.xml': '<b x="1"/>', 'd.xml': '<d/>' });
		expect(await diffOpcPackages(base, other)).toEqual({
			added: ['d.xml'],
			removed: ['c.xml'],
			changed: ['b.xml'],
		});
	});

	it('ignores compression and part order', async () => {
		const base = await pack({ 'a.xml': '<a/>', 'b.xml': '<b/>' }, 'DEFLATE');
		const other = await pack({ 'b.xml': '<b/>', 'a.xml': '<a/>' }, 'STORE');
		expect(await diffOpcPackages(base, other)).toEqual({ added: [], removed: [], changed: [] });
	});
});
