import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { mergeDocxSnapshots } from './merge-snapshots';
import { loadDocx } from './parse';
import { saveDocx } from './save';
import { createDocument } from './model';
import { NS } from '../xml/namespaces';
import type { SnapshotMergeResult } from '../opc/merge-snapshots';

const paragraph = (text: string, id = 'A', properties = '') =>
	`<w:p w14:paraId="${id}"><w:r>${properties}<w:t>${text}</w:t></w:r></w:p>`;
const contentType =
	'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml';
async function docx(body: string, extra: Record<string, string | Uint8Array> = {}) {
	const zip = new JSZip();
	zip.file(
		'[Content_Types].xml',
		`<Types xmlns="${NS.contentTypes}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="bin" ContentType="application/octet-stream"/><Override PartName="/word/document.xml" ContentType="${contentType}"/></Types>`,
	);
	zip.file(
		'_rels/.rels',
		`<Relationships xmlns="${NS.rels}"><Relationship Id="rId1" Type="${NS.r}/officeDocument" Target="word/document.xml"/></Relationships>`,
	);
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${NS.w}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:r="${NS.r}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
	);
	for (const [name, bytes] of Object.entries(extra)) zip.file(name, bytes);
	return zip.generateAsync({ type: 'uint8array' });
}
async function mergedZip(result: SnapshotMergeResult) {
	expect(result.status, result.status === 'unresolved' ? result.issues.join('\n') : '').toBe(
		'merged',
	);
	if (result.status !== 'merged') throw new Error('Expected merged DOCX');
	return JSZip.loadAsync(result.bytes);
}
async function mergedXml(base: Uint8Array, local: Uint8Array, remote: Uint8Array) {
	return (await mergedZip(await mergeDocxSnapshots(base, local, remote)))
		.file('word/document.xml')!
		.async('string');
}

describe('mergeDocxSnapshots', () => {
	it('retains concurrent same-position insertions in a deterministic order', async () => {
		const base = await docx(paragraph('Hello'));
		const local = await docx(paragraph('AHello'));
		const remote = await docx(paragraph('BHello'));
		expect(await mergedXml(base, local, remote)).toContain('ABHello');
		expect(await mergedXml(base, remote, local)).toContain('ABHello');
	});

	it('combines non-overlapping edits to one text run and produces a loadable document', async () => {
		const base = await docx(paragraph('Hello world'));
		const result = await mergeDocxSnapshots(
			base,
			await docx(paragraph('Dear Hello world')),
			await docx(paragraph('Hello world!')),
		);
		const zip = await mergedZip(result);
		expect(await zip.file('word/document.xml')!.async('string')).toContain('Dear Hello world!');
		if (result.status === 'merged') {
			const loaded = await loadDocx(result.bytes);
			expect(loaded.model.blocks.length).toBeGreaterThan(0);
		}
	});

	it('combines edits to different identified paragraphs without changing their order', async () => {
		const base = await docx(paragraph('Alpha') + paragraph('Beta', 'B'));
		const xml = await mergedXml(
			base,
			await docx(paragraph('Alpha local') + paragraph('Beta', 'B')),
			await docx(paragraph('Alpha') + paragraph('Beta remote', 'B')),
		);
		expect(xml).toContain('Alpha local');
		expect(xml).toContain('Beta remote');
		expect(xml.indexOf('Alpha local')).toBeLessThan(xml.indexOf('Beta remote'));
	});

	it('combines aligned paragraphs without IDs while refusing a concurrent reorder', async () => {
		const p = (text: string) => paragraph(text).replace(' w14:paraId="A"', '');
		const base = await docx(p('Alpha') + p('Beta'));
		const local = await docx(p('Alpha local') + p('Beta'));
		const remote = await docx(p('Alpha') + p('Beta remote'));
		expect(await mergedXml(base, local, remote)).toContain('Beta remote');
		const reordered = await mergeDocxSnapshots(base, await docx(p('Beta') + p('Alpha')), remote);
		expect(reordered.status).toBe('unresolved');
	});

	it('keeps independent run formatting and text edits on the same run', async () => {
		const base = await docx(paragraph('Heading', 'A', '<w:rPr><w:b w:val="0"/></w:rPr>'));
		const xml = await mergedXml(
			base,
			await docx(paragraph('Heading', 'A', '<w:rPr><w:b w:val="1"/></w:rPr>')),
			await docx(paragraph('New Heading', 'A', '<w:rPr><w:b w:val="0"/></w:rPr>')),
		);
		expect(xml).toContain('w:val="1"');
		expect(xml).toContain('New Heading');
	});

	it('preserves independent opaque parts and one-sided structural additions', async () => {
		const media = new Uint8Array([1, 2, 3, 255]);
		const base = await docx(paragraph('Base'), { 'word/media/opaque.bin': new Uint8Array([0]) });
		const result = await mergeDocxSnapshots(
			base,
			await docx(paragraph('Base') + paragraph('Added', 'B'), { 'word/media/opaque.bin': media }),
			await docx(paragraph('Base remote'), { 'word/media/opaque.bin': new Uint8Array([0]) }),
		);
		const zip = await mergedZip(result);
		expect(await zip.file('word/media/opaque.bin')!.async('uint8array')).toEqual(media);
		const xml = await zip.file('word/document.xml')!.async('string');
		expect(xml).toContain('Added');
		expect(xml).toContain('Base remote');
	});

	it('retains unknown XML extensions inside a paragraph being merged', async () => {
		const decorate = (text: string) =>
			paragraph(text).replace(
				'</w:p>',
				'<custom:payload xmlns:custom="urn:merge-test" value="opaque"><custom:data>untouched</custom:data></custom:payload></w:p>',
			);
		const xml = await mergedXml(
			await docx(decorate('Hello')),
			await docx(decorate('Dear Hello')),
			await docx(decorate('Hello!')),
		);
		expect(xml).toContain('Dear Hello!');
		expect(xml).toContain('value="opaque"');
		expect(xml).toContain('untouched');
	});

	it('does not merge overlapping text replacements or deletion versus editing', async () => {
		const base = await docx(paragraph('Original'));
		expect(
			(
				await mergeDocxSnapshots(
					base,
					await docx(paragraph('Local')),
					await docx(paragraph('Remote')),
				)
			).status,
		).toBe('unresolved');
		expect(
			(await mergeDocxSnapshots(base, await docx(''), await docx(paragraph('Remote')))).status,
		).toBe('unresolved');
	});

	it('does not invent run alignment or silently lose concurrent opaque edits', async () => {
		const base = await docx(paragraph('Original'), {
			'word/media/opaque.bin': new Uint8Array([0]),
		});
		const local = await docx(
			'<w:p w14:paraId="A"><w:r><w:t>Ori</w:t></w:r><w:r><w:t>ginal</w:t></w:r></w:p>',
			{ 'word/media/opaque.bin': new Uint8Array([0]) },
		);
		expect(
			(
				await mergeDocxSnapshots(
					base,
					local,
					await docx(paragraph('Original remote'), {
						'word/media/opaque.bin': new Uint8Array([0]),
					}),
				)
			).status,
		).toBe('unresolved');
		expect(
			(
				await mergeDocxSnapshots(
					base,
					await docx(paragraph('Original'), { 'word/media/opaque.bin': new Uint8Array([1]) }),
					await docx(paragraph('Original'), { 'word/media/opaque.bin': new Uint8Array([2]) }),
				)
			).status,
		).toBe('unresolved');
	});

	it('refuses collisions between independently added paragraph identities', async () => {
		const base = await docx(paragraph('Base'));
		expect(
			(
				await mergeDocxSnapshots(
					base,
					await docx(paragraph('Base') + paragraph('Local new', 'B')),
					await docx(paragraph('Base') + paragraph('Remote new', 'B')),
				)
			).status,
		).toBe('unresolved');
	});

	it('merges actual DOCX-writer output with different last-modified metadata', async () => {
		const model = createDocument();
		model.blocks = [{ type: 'paragraph', id: 'p1', runs: [{ text: 'Hello world' }] }];
		model.properties = { title: 'Merge fixture', lastModifiedBy: 'Base author' };
		const base = await saveDocx(model);
		const edit = async (text: string, author: string, at: string) => {
			const loaded = await loadDocx(base);
			const first = loaded.model.blocks[0];
			if (first?.type !== 'paragraph' || !first.runs[0])
				throw new Error('Missing fixture paragraph');
			first.runs[0].text = text;
			return loaded.save(loaded.model, undefined, { modified: { author, at: new Date(at) } });
		};
		const result = await mergeDocxSnapshots(
			base,
			await edit('Dear Hello world', 'Local author', '2026-11-04T00:00:00Z'),
			await edit('Hello world!', 'Remote author', '2026-11-04T00:01:00Z'),
		);
		const zip = await mergedZip(result);
		expect(await zip.file('word/document.xml')!.async('string')).toContain('Dear Hello world!');
		if (result.status === 'merged') {
			const loaded = await loadDocx(result.bytes);
			const first = loaded.model.blocks[0];
			expect(first?.type === 'paragraph' && first.runs.map((run) => run.text).join('')).toBe(
				'Dear Hello world!',
			);
		}
	});
});
