import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { mergeOpcSnapshots } from './merge-snapshots';
async function zip(parts: Record<string, string | Uint8Array>) {
	const z = new JSZip();
	for (const [name, value] of Object.entries(parts)) z.file(name, value);
	return z.generateAsync({ type: 'uint8array' });
}
describe('three-way OPC merge', () => {
	it('preserves independent opaque parts and binary bytes', async () => {
		const result = await mergeOpcSnapshots(
			await zip({ a: 'A', b: 'B', image: new Uint8Array([1, 2]) }),
			await zip({ a: 'AA', b: 'B', image: new Uint8Array([1, 2]) }),
			await zip({ a: 'A', b: 'BB', image: new Uint8Array([1, 2]) }),
		);
		expect(result.status).toBe('merged');
		if (result.status === 'merged') {
			const merged = await JSZip.loadAsync(result.bytes);
			expect(await merged.file('a')!.async('string')).toBe('AA');
			expect(await merged.file('b')!.async('string')).toBe('BB');
			expect(await merged.file('image')!.async('uint8array')).toEqual(new Uint8Array([1, 2]));
		}
	});
	it('declines divergent binary changes and same-name additions', async () => {
		expect(
			(await mergeOpcSnapshots(await zip({ a: 'A' }), await zip({ a: 'B' }), await zip({ a: 'C' })))
				.status,
		).toBe('unresolved');
		expect(
			(await mergeOpcSnapshots(await zip({}), await zip({ a: 'B' }), await zip({ a: 'C' }))).status,
		).toBe('unresolved');
	});
	it('combines independently added relationship ids', async () => {
		const rels = (value: string) =>
			`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${value}</Relationships>`;
		const result = await mergeOpcSnapshots(
			await zip({ '_rels/.rels': rels('') }),
			await zip({ '_rels/.rels': rels('<Relationship Id="r1" Target="a" Type="t"/>'), a: 'A' }),
			await zip({ '_rels/.rels': rels('<Relationship Id="r2" Target="b" Type="t"/>'), b: 'B' }),
		);
		expect(result.status).toBe('merged');
	});
	it('merges different core properties added by each peer, in either order', async () => {
		const core = (value: string) =>
			`<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>A</dc:creator>${value}</cp:coreProperties>`;
		const title = core('<dc:title>T</dc:title>');
		const subject = core('<dc:subject>S</dc:subject>');
		for (const [local, remote] of [
			[title, subject],
			[subject, title],
		]) {
			const result = await mergeOpcSnapshots(
				await zip({ 'docProps/core.xml': core('') }),
				await zip({ 'docProps/core.xml': local! }),
				await zip({ 'docProps/core.xml': remote! }),
			);
			expect(result.status).toBe('merged');
			if (result.status !== 'merged') continue;
			const merged = await JSZip.loadAsync(result.bytes);
			const xml = await merged.file('docProps/core.xml')!.async('string');
			expect(xml).toContain('<dc:title>T</dc:title>');
			expect(xml).toContain('<dc:subject>S</dc:subject>');
		}
		expect(
			(
				await mergeOpcSnapshots(
					await zip({ 'docProps/core.xml': core('') }),
					await zip({ 'docProps/core.xml': core('<dc:title>A</dc:title>') }),
					await zip({ 'docProps/core.xml': core('<dc:title>B</dc:title>') }),
				)
			).status,
		).toBe('unresolved');
	});
	it('merges newly added volatile timestamps but rejects dangling relationship targets', async () => {
		const core = (modified: string) =>
			`<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dcterms="http://purl.org/dc/terms/">${modified ? `<dcterms:modified>${modified}</dcterms:modified>` : ''}</cp:coreProperties>`;
		const result = await mergeOpcSnapshots(
			await zip({ 'docProps/core.xml': core('') }),
			await zip({ 'docProps/core.xml': core('2026-11-04T01:00:00Z') }),
			await zip({ 'docProps/core.xml': core('2026-11-04T02:00:00Z') }),
		);
		expect(result.status).toBe('merged');
		const rel = '<Relationships><Relationship Id="r1" Target="gone.xml" Type="t"/></Relationships>';
		expect(
			(
				await mergeOpcSnapshots(
					await zip({ a: '1' }),
					await zip({ a: '2' }),
					await zip({ a: '1', '_rels/.rels': rel }),
				)
			).status,
		).toBe('unresolved');
	});
	it('surfaces callback failures and invalid packages without returning bytes', async () => {
		expect(
			(
				await mergeOpcSnapshots(
					await zip({ a: 'A' }),
					await zip({ a: 'B' }),
					await zip({ a: 'C' }),
					{
						mergePart: () => {
							throw new Error('private');
						},
					},
				)
			).status,
		).toBe('unresolved');
		expect(
			(await mergeOpcSnapshots(new Uint8Array([1]), new Uint8Array([2]), new Uint8Array([3])))
				.status,
		).toBe('unresolved');
	});
});
