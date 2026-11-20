import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { clonePart } from './clone-part-graph';
import { findSharedOwnedParts } from './part-ownership';
import { buildRelationshipsXml, parseRelationships, relationshipsPartFor } from './relationships';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const rel = (type: string, target: string) => ({
	type: `${R}/${type}`,
	target,
	mode: 'Internal' as const,
});

function slide(frames: string[]): string {
	const body = frames
		.map((id) => `<p:graphicFrame><c:chart xmlns:c="c" r:id="${id}"/></p:graphicFrame>`)
		.join('');
	return `<p:sld xmlns:p="p" xmlns:r="${R}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">${body}</p:sld>`;
}

async function deck(
	parts: Record<string, string>,
	rels: Record<string, Record<string, ReturnType<typeof rel>>>,
) {
	const zip = new JSZip();
	for (const [name, xml] of Object.entries(parts)) zip.file(name, xml);
	for (const [source, map] of Object.entries(rels))
		zip.file(relationshipsPartFor(source), buildRelationshipsXml(new Map(Object.entries(map))));
	return zip;
}

describe('findSharedOwnedParts', () => {
	it('counts each frame citing a chart relationship as an owner', async () => {
		const zip = await deck(
			{ 'ppt/slides/slide1.xml': slide(['rId2', 'rId2']), 'ppt/charts/chart1.xml': '<c/>' },
			{ 'ppt/slides/slide1.xml': { rId2: rel('chart', '../charts/chart1.xml') } },
		);
		expect(await findSharedOwnedParts(zip)).toStrictEqual([
			{
				part: 'ppt/charts/chart1.xml',
				claims: [
					{ source: 'ppt/slides/slide1.xml', relationshipId: 'rId2' },
					{ source: 'ppt/slides/slide1.xml', relationshipId: 'rId2' },
				],
			},
		]);
	});

	it('finds two slides owning one chart, and ignores a relationship nothing cites', async () => {
		const shared = await deck(
			{
				'ppt/slides/slide1.xml': slide(['rId2']),
				'ppt/slides/slide2.xml': slide(['rId5']),
				'ppt/charts/chart1.xml': '<c/>',
			},
			{
				'ppt/slides/slide1.xml': { rId2: rel('chart', '../charts/chart1.xml') },
				'ppt/slides/slide2.xml': { rId5: rel('chart', '../charts/chart1.xml') },
			},
		);
		expect((await findSharedOwnedParts(shared)).map((item) => item.part)).toStrictEqual([
			'ppt/charts/chart1.xml',
		]);
		// PowerPoint opens a package whose slide keeps a relationship a moved chart left behind.
		const stale = await deck(
			{
				'ppt/slides/slide1.xml': slide(['rId2']),
				'ppt/slides/slide2.xml': slide([]),
				'ppt/charts/chart1.xml': '<c/>',
			},
			{
				'ppt/slides/slide1.xml': { rId2: rel('chart', '../charts/chart1.xml') },
				'ppt/slides/slide2.xml': { rId9: rel('chart', '../charts/chart1.xml') },
			},
		);
		expect(await findSharedOwnedParts(stale)).toStrictEqual([]);
	});

	it('reads an mc:Fallback as the same element, and leaves images shared', async () => {
		const xml = `<p:sld xmlns:p="p" xmlns:r="${R}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:AlternateContent><mc:Choice><p:oleObj r:id="rId2"/></mc:Choice><mc:Fallback><p:oleObj r:id="rId2"/></mc:Fallback></mc:AlternateContent><p:pic r:embed="rId3"/><p:pic r:embed="rId3"/></p:sld>`;
		const zip = await deck(
			{ 'ppt/slides/slide1.xml': xml },
			{
				'ppt/slides/slide1.xml': {
					rId2: rel('package', '../embeddings/book.xlsx'),
					rId3: rel('image', '../media/image1.png'),
				},
			},
		);
		expect(await findSharedOwnedParts(zip)).toStrictEqual([]);
	});

	it('finds two masters on one theme, but not the presentation pointing at it', async () => {
		const zip = await deck(
			{
				'ppt/presentation.xml': '<p:presentation xmlns:p="p"/>',
				'ppt/slideMasters/slideMaster1.xml': '<p:sldMaster xmlns:p="p"/>',
				'ppt/slideMasters/slideMaster2.xml': '<p:sldMaster xmlns:p="p"/>',
			},
			{
				'ppt/presentation.xml': { rId1: rel('theme', 'theme/theme1.xml') },
				'ppt/slideMasters/slideMaster1.xml': { rId1: rel('theme', '../theme/theme1.xml') },
				'ppt/slideMasters/slideMaster2.xml': { rId1: rel('theme', '../theme/theme1.xml') },
			},
		);
		expect(await findSharedOwnedParts(zip)).toStrictEqual([
			{
				part: 'ppt/theme/theme1.xml',
				claims: [
					{ source: 'ppt/slideMasters/slideMaster1.xml', relationshipId: 'rId1' },
					{ source: 'ppt/slideMasters/slideMaster2.xml', relationshipId: 'rId1' },
				],
			},
		]);
	});
});

describe('clonePart', () => {
	it('copies a part under a free name with its content type and owned parts', async () => {
		const zip = await deck(
			{
				'[Content_Types].xml':
					'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/ppt/charts/chart1.xml" ContentType="chart/type"/></Types>',
				'ppt/charts/chart1.xml': '<chart/>',
				'ppt/charts/chart2.xml': '<other/>',
				'ppt/embeddings/book1.xlsx': 'bytes',
			},
			{ 'ppt/charts/chart1.xml': { rId1: rel('package', '../embeddings/book1.xlsx') } },
		);
		const copies = await clonePart({ zip, part: 'ppt/charts/chart1.xml', shouldClone: () => true });
		expect(Object.fromEntries(copies)).toStrictEqual({
			'ppt/charts/chart1.xml': 'ppt/charts/chart3.xml',
			'ppt/embeddings/book1.xlsx': 'ppt/embeddings/book2.xlsx',
		});
		expect(await zip.file('ppt/charts/chart3.xml')!.async('string')).toBe('<chart/>');
		expect(await zip.file('[Content_Types].xml')!.async('string')).toContain(
			'<Override PartName="/ppt/charts/chart3.xml" ContentType="chart/type"/>',
		);
		const rels = parseRelationships(
			await zip.file(relationshipsPartFor('ppt/charts/chart3.xml'))!.async('string'),
		);
		expect(rels.get('rId1')?.target).toBe('../embeddings/book2.xlsx');
	});
});
