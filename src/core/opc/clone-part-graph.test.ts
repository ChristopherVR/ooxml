import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { clonePartGraph, relativePartTarget } from './clone-part-graph';
import { buildContentTypesXml, contentTypeForPart, parseContentTypes } from './content-types';
import {
	buildRelationshipsXml,
	parseRelationships,
	relationshipsPartFor,
	type RelationshipInput,
} from './relationships';

const rel = (target: string, type = 'owned') => ({ type, target, mode: 'Internal' as const });

describe('clonePartGraph', () => {
	it('copies owned dependencies once, remaps cycles and keeps shared and external targets', async () => {
		const zip = new JSZip();
		zip.file('doc/source.xml', '<source/>');
		zip.file('parts/data7.xml', '<data/>');
		zip.file('parts/book1.bin', new Uint8Array([1, 2, 3]));
		zip.file('parts/shared.xml', '<shared/>');
		zip.file(
			relationshipsPartFor('doc/source.xml'),
			buildRelationshipsXml(
				new Map<string, RelationshipInput>([
					['customChart', rel('../parts/data7.xml')],
					['sameChart', rel('../parts/data7.xml')],
					['shared', rel('../parts/shared.xml', 'shared')],
					[
						'url',
						{
							type: 'hyperlink',
							target: 'https://example.com/?x=1&y=2',
							mode: 'External' as const,
						},
					],
				]),
			),
		);
		zip.file(
			relationshipsPartFor('parts/data7.xml'),
			buildRelationshipsXml(
				new Map([
					['back', rel('../doc/source.xml')],
					['book', rel('book1.bin')],
				]),
			),
		);
		zip.file(
			'[Content_Types].xml',
			buildContentTypesXml({
				defaults: new Map([['bin', 'application/octet-stream']]),
				overrides: new Map([
					['/doc/source.xml', 'root/type'],
					['/parts/data7.xml', 'data/type'],
				]),
			}),
		);
		const copies = await clonePartGraph({
			zip,
			sourcePart: 'doc/source.xml',
			targetPart: 'other/copy.xml',
			shouldClone: (relationship) => relationship.type !== 'shared',
		});
		expect(copies.size).toBe(3);
		const dataCopy = copies.get('parts/data7.xml')!;
		const bookCopy = copies.get('parts/book1.bin')!;
		expect(dataCopy).not.toBe('parts/data7.xml');
		expect(await zip.file(dataCopy)!.async('string')).toBe('<data/>');
		expect(await zip.file(bookCopy)!.async('uint8array')).toEqual(new Uint8Array([1, 2, 3]));
		const rootRels = parseRelationships(
			await zip.file(relationshipsPartFor('other/copy.xml'))!.async('string'),
		);
		expect(rootRels.get('customChart')?.target).toBe(`../${dataCopy}`);
		expect(rootRels.get('sameChart')?.target).toBe(`../${dataCopy}`);
		expect(rootRels.get('shared')?.target).toBe('../parts/shared.xml');
		expect(rootRels.get('url')?.target).toBe('https://example.com/?x=1&y=2');
		const dataRels = parseRelationships(
			await zip.file(relationshipsPartFor(dataCopy))!.async('string'),
		);
		expect(dataRels.get('back')?.target).toBe('../other/copy.xml');
		const types = parseContentTypes(await zip.file('[Content_Types].xml')!.async('string'));
		expect(contentTypeForPart(types, 'other/copy.xml')).toBe('root/type');
		expect(contentTypeForPart(types, dataCopy)).toBe('data/type');
		expect(contentTypeForPart(types, bookCopy)).toBe('application/octet-stream');
	});

	it('reports missing owned parts rather than emitting a dangling copy', async () => {
		const zip = new JSZip();
		zip.file(
			relationshipsPartFor('source.xml'),
			buildRelationshipsXml(new Map([['rId1', rel('missing.xml')]])),
		);
		await expect(
			clonePartGraph({
				zip,
				sourcePart: 'source.xml',
				targetPart: 'copy.xml',
				shouldClone: () => true,
			}),
		).rejects.toThrow('missing.xml');
	});

	it('creates relative targets across package folders', () => {
		expect(relativePartTarget('ppt/slides/slide2.xml', 'ppt/charts/chart2.xml')).toBe(
			'../charts/chart2.xml',
		);
		expect(relativePartTarget('ppt/charts/chart2.xml', 'ppt/charts/style2.xml')).toBe('style2.xml');
		expect(relativePartTarget('root.xml', 'other.xml')).toBe('other.xml');
	});
});

it('imports a separate archive with shared cyclic infrastructure and conflicting content types', async () => {
	const source = new JSZip(),
		zip = new JSZip();
	source.file(
		'[Content_Types].xml',
		buildContentTypesXml({
			defaults: new Map([['bin', 'source/binary']]),
			overrides: new Map([
				['/roots/a.xml', 'source/root'],
				['/roots/b.xml', 'source/root'],
				['/parts/master1.xml', 'source/master'],
			]),
		}),
	);
	zip.file(
		'[Content_Types].xml',
		buildContentTypesXml({
			defaults: new Map([['bin', 'destination/binary']]),
			overrides: new Map(),
		}),
	);
	source.file('roots/a.xml', '<a/>');
	source.file('roots/b.xml', '<b/>');
	source.file('parts/master1.xml', '<master/>');
	source.file('parts/layout1.bin', new Uint8Array([5, 7]));
	zip.file('parts/master1.xml', '<destination/>');
	zip.file('parts/layout1.bin', new Uint8Array([9]));
	for (const root of ['a', 'b'])
		source.file(
			relationshipsPartFor(`roots/${root}.xml`),
			buildRelationshipsXml(
				new Map<string, RelationshipInput>([
					['nativeId', rel('../parts/master1.xml')],
					[
						'link',
						{ type: 'hyperlink', target: 'https://example.com/path?x=1&y=2', mode: 'External' },
					],
				]),
			),
		);
	source.file(
		relationshipsPartFor('parts/master1.xml'),
		buildRelationshipsXml(new Map([['layout', rel('layout1.bin')]])),
	);
	source.file(
		relationshipsPartFor('parts/layout1.bin'),
		buildRelationshipsXml(
			new Map([
				['master', rel('master1.xml')],
				['slide', rel('../roots/b.xml')],
			]),
		),
	);
	const copies = new Map([
		['roots/a.xml', 'roots/copy1.xml'],
		['roots/b.xml', 'roots/copy2.xml'],
	]);
	for (const root of ['a', 'b'])
		await clonePartGraph({
			zip,
			sourceZip: source,
			copies,
			sourcePart: `roots/${root}.xml`,
			targetPart: copies.get(`roots/${root}.xml`)!,
			shouldClone: () => true,
		});
	expect(copies.size).toBe(4);
	const master = copies.get('parts/master1.xml')!,
		layout = copies.get('parts/layout1.bin')!;
	expect(master).not.toBe('parts/master1.xml');
	expect(layout).not.toBe('parts/layout1.bin');
	expect(await zip.file('parts/master1.xml')!.async('string')).toBe('<destination/>');
	expect(await source.file('parts/master1.xml')!.async('string')).toBe('<master/>');
	expect(await zip.file(layout)!.async('uint8array')).toEqual(new Uint8Array([5, 7]));
	const types = parseContentTypes(await zip.file('[Content_Types].xml')!.async('string'));
	expect(contentTypeForPart(types, layout)).toBe('source/binary');
	expect(contentTypeForPart(types, 'parts/layout1.bin')).toBe('destination/binary');
	const layoutRels = parseRelationships(
		await zip.file(relationshipsPartFor(layout))!.async('string'),
	);
	expect(layoutRels.get('slide')?.target).toBe('../roots/copy2.xml');
	for (const root of ['roots/copy1.xml', 'roots/copy2.xml']) {
		const rels = parseRelationships(await zip.file(relationshipsPartFor(root))!.async('string'));
		expect(rels.get('nativeId')?.target).toBe(`../${master}`);
		expect(rels.get('link')).toMatchObject({
			target: 'https://example.com/path?x=1&y=2',
			mode: 'External',
		});
	}
});

it('rejects dependencies absent from the source even when that name exists in the destination', async () => {
	const zip = new JSZip(),
		source = new JSZip();
	zip.file('missing.xml', '<foreign/>');
	source.file(
		relationshipsPartFor('source.xml'),
		buildRelationshipsXml(new Map([['rId1', rel('missing.xml')]])),
	);
	await expect(
		clonePartGraph({
			zip,
			sourceZip: source,
			sourcePart: 'source.xml',
			targetPart: 'copied.xml',
			shouldClone: () => true,
		}),
	).rejects.toThrow('missing.xml');
});
