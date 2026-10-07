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
