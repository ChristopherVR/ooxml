// Reusing the styles, themes and masters of an earlier read must give the document a whole read
// gives, and must read them afresh whenever a part they came from changed.
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { addMaster, masterShape } from './collab/test-support/masters';
import {
	type VisioDocumentCache,
	createVsdx,
	editVsdx,
	mergeVisioPages,
	parseVsdx,
	parseVsdxPages,
} from './index';

const rect = (pageId: string, shapeId: string, x: number) =>
	({ type: 'create-rectangle', pageId, shapeId, x, y: 2, width: 1, height: 1 }) as const;

async function withMaster(): Promise<Uint8Array> {
	let bytes = (await editVsdx(await createVsdx(), [rect('0', '1', 1), rect('0', '2', 3)])).bytes;
	bytes = (
		await editVsdx(bytes, [{ type: 'insert-page', pageId: '5', afterPageId: '0', name: 'Two' }])
	).bytes;
	bytes = (await editVsdx(bytes, [rect('5', '1', 2)])).bytes;
	const zip = await JSZip.loadAsync(bytes);
	await addMaster(zip, { id: '2', unique: '{A}', shapeId: '77', shapes: masterShape('1', 1) });
	return zip.generateAsync({ type: 'uint8array' });
}
async function patched(bytes: Uint8Array, path: string, change: (text: string) => string) {
	const zip = await JSZip.loadAsync(bytes);
	zip.file(path, change((await zip.file(path)!.async('string')) as string));
	return zip.generateAsync({ type: 'uint8array' });
}

describe('the document cache', () => {
	it('gives the document a whole read gives, read after read', async () => {
		let bytes = await withMaster();
		const cache: VisioDocumentCache = {};
		let previous = await parseVsdx(bytes, { documentCache: cache });
		expect(previous).toEqual(await parseVsdx(bytes));
		expect(previous.masters?.length).toBeGreaterThan(0);
		const entry = cache.entry;
		expect(entry).toBeDefined();
		for (const x of [6, 7, 8]) {
			bytes = (await editVsdx(bytes, [{ type: 'move-shape', pageId: '5', shapeId: '1', x, y: 2 }]))
				.bytes;
			const partial = await parseVsdx(bytes, { pages: ['5'], documentCache: cache });
			// Nothing document-level changed: the entry was reused, not replaced.
			expect(cache.entry).toBe(entry);
			const merged = mergeVisioPages(previous, partial, ['5']);
			expect(merged).toEqual(await parseVsdx(bytes));
			previous = merged!;
		}
		// A whole read through the cache is the same too.
		expect(await parseVsdx(bytes, { documentCache: cache })).toEqual(await parseVsdx(bytes));
		expect(cache.entry).toBe(entry);
	});

	it('merges to a whole read through parseVsdxPages with a cache', async () => {
		const before = await withMaster();
		const cache: VisioDocumentCache = {};
		const previous = await parseVsdx(before, { documentCache: cache });
		const after = (
			await editVsdx(before, [
				rect('0', '9', 5),
				{ type: 'replace-plain-text', pageId: '5', shapeId: '1', text: 'hi' },
			])
		).bytes;
		const merged = await parseVsdxPages(after, previous, ['0', '5'], { documentCache: cache });
		expect(merged).toEqual(await parseVsdx(after));
	});

	it('reads afresh when a master, the document part or an option changed', async () => {
		const bytes = await withMaster();
		const cache: VisioDocumentCache = {};
		await parseVsdx(bytes, { pages: ['0'], documentCache: cache });
		const first = cache.entry;
		const master = Object.keys((await JSZip.loadAsync(bytes)).files).find((name) =>
			/^visio\/masters\/master\d+\.xml$/.test(name),
		)!;
		const newMaster = await patched(bytes, master, (t) =>
			t.replace('</MasterContents>', '<!-- x --></MasterContents>'),
		);
		const doc1 = await parseVsdx(newMaster, { pages: ['0'], documentCache: cache });
		expect(cache.entry).not.toBe(first);
		expect(doc1.masters).toEqual((await parseVsdx(newMaster, { pages: ['0'] })).masters);
		const second = cache.entry;
		const newDoc = await patched(newMaster, 'visio/document.xml', (t) =>
			t.replace('</VisioDocument>', '<!-- y --></VisioDocument>'),
		);
		await parseVsdx(newDoc, { pages: ['0'], documentCache: cache });
		expect(cache.entry).not.toBe(second);
		const third = cache.entry;
		await parseVsdx(newDoc, { pages: ['0'], documentCache: cache, maxShapes: 10_000 });
		expect(cache.entry).not.toBe(third);
	});

	it('keeps a limit consumed by the document part counting for the pages', async () => {
		const bytes = await withMaster();
		const options = { pages: ['0'], maxGeometryCommands: 3 } as const;
		const cache: VisioDocumentCache = {};
		const outcome = async (documentCache?: VisioDocumentCache) =>
			parseVsdx(bytes, { ...options, ...(documentCache ? { documentCache } : {}) }).then(
				(d) => d.diagnostics.length,
				(e: Error) => e.message,
			);
		const plain = await outcome();
		expect(await outcome(cache)).toEqual(plain); // populates
		expect(await outcome(cache)).toEqual(plain); // reuses
	});
});
