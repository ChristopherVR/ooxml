import { XMLParser } from 'fast-xml-parser';
import JSZip from 'jszip';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { guardZip, ZipLimitError } from '../../../../opc/zip-limits';
import type { XmlObject } from '../../types';
import { LoadPartReader, readPartText, readPartXml } from './load-part-reader';

const relsPath = 'ppt/slides/_rels/slide1.xml.rels';
const relsXml =
	'<Relationships><Relationship Id="rId1" Type="layout" Target="../slideLayouts/layout1.xml"><Extra value="original"/></Relationship></Relationships>';

function setup() {
	const zip = new JSZip().file(relsPath, relsXml).file('other.xml', '<root value="old"/>');
	const parser = new XMLParser({ ignoreAttributes: false });
	const reads = vi.spyOn(zip.file(relsPath)!, 'async');
	const parses = vi.spyOn(parser, 'parse');
	const reader = new LoadPartReader(zip);
	return { zip, parser, reads, parses, reader };
}

function relationship(data: XmlObject | undefined): XmlObject {
	return (data!.Relationships as XmlObject).Relationship as XmlObject;
}

afterEach(() => {
	vi.restoreAllMocks();
});

describe('load part reader', () => {
	it('shares concurrent text reads and parses each relationship part once', async () => {
		const { zip, parser, reads, parses, reader } = setup();
		const values = await Promise.all([
			readPartText(zip, relsPath),
			readPartXml(zip, parser, relsPath),
			readPartXml(zip, parser, relsPath),
		]);
		expect(values[0]).toBe(relsXml);
		expect(values[1]).toEqual(values[2]);
		expect(values[1]).not.toBe(values[2]);
		expect(reads).toHaveBeenCalledTimes(1);
		expect(parses).toHaveBeenCalledTimes(1);
		reader.close();
	});

	it('returns independent copies of relationship trees', async () => {
		const { zip, parser, reader } = setup();
		const first = await readPartXml(zip, parser, relsPath);
		relationship(first)['@_Target'] = 'changed.xml';
		(relationship(first).Extra as XmlObject)['@_value'] = 'changed';
		const second = await readPartXml(zip, parser, relsPath);
		expect(relationship(second)['@_Target']).toBe('../slideLayouts/layout1.xml');
		expect((relationship(second).Extra as XmlObject)['@_value']).toBe('original');
		reader.close();
	});

	it('parses relationship parts once per parser', async () => {
		const { zip, parser, reader } = setup();
		const other = new XMLParser({ ignoreAttributes: true });
		const otherParses = vi.spyOn(other, 'parse');
		await readPartXml(zip, parser, relsPath);
		expect(relationship(await readPartXml(zip, other, relsPath))['@_Target']).toBeUndefined();
		await readPartXml(zip, other, relsPath);
		expect(otherParses).toHaveBeenCalledTimes(1);
		reader.close();
	});

	it('observes parts written, replaced or removed while open', async () => {
		const { zip, parser, reads, reader } = setup();
		await readPartXml(zip, parser, relsPath);
		zip.file('other.xml', '<root/>');
		await readPartXml(zip, parser, relsPath);
		expect(reads).toHaveBeenCalledTimes(1);
		zip.file(relsPath, relsXml.replace('layout1.xml', 'layout2.xml'));
		expect(relationship(await readPartXml(zip, parser, relsPath))['@_Target']).toBe(
			'../slideLayouts/layout2.xml',
		);
		zip.remove(relsPath);
		expect(await readPartText(zip, relsPath)).toBeUndefined();
		expect(await readPartXml(zip, parser, relsPath)).toBeUndefined();
		zip.file(relsPath, relsXml);
		expect(await readPartText(zip, relsPath)).toBe(relsXml);
		reader.close();
	});

	it('retries rejected reads and rejected relationship parses', async () => {
		const { zip, parser, reads, parses, reader } = setup();
		reads.mockRejectedValueOnce(new Error('read failure'));
		await expect(readPartText(zip, relsPath)).rejects.toThrow('read failure');
		parses.mockImplementationOnce(() => {
			throw new Error('parse failure');
		});
		await expect(readPartXml(zip, parser, relsPath)).rejects.toThrow('parse failure');
		expect(await readPartXml(zip, parser, relsPath)).toBeDefined();
		expect(reads).toHaveBeenCalledTimes(2);
		expect(parses).toHaveBeenCalledTimes(2);
		reader.close();
	});

	it('shares the text of other parts but parses them per call', async () => {
		const { zip, parser, parses, reader } = setup();
		const reads = vi.spyOn(zip.file('other.xml')!, 'async');
		const first = await readPartXml(zip, parser, 'other.xml');
		(first!.root as XmlObject)['@_value'] = 'changed';
		const second = await readPartXml(zip, parser, 'other.xml');
		expect((second!.root as XmlObject)['@_value']).toBe('old');
		expect(reads).toHaveBeenCalledTimes(1);
		expect(parses).toHaveBeenCalledTimes(2);
		reader.close();
	});

	it('does not retain a missing or empty relationship part', async () => {
		const { zip, parser, reader } = setup();
		expect(await readPartXml(zip, parser, 'missing.rels')).toBeUndefined();
		zip.file('missing.rels', '');
		expect(await readPartXml(zip, parser, 'missing.rels')).toBeUndefined();
		zip.file('missing.rels', '<Relationships/>');
		expect(await readPartXml(zip, parser, 'missing.rels')).toEqual({ Relationships: '' });
		reader.close();
	});

	it('reads the archive directly once closed, and settles reads already started', async () => {
		const { zip, parser, reads, parses, reader } = setup();
		const pending = readPartXml(zip, parser, relsPath);
		reader.close();
		expect(reader.isOpen).toBe(false);
		expect(relationship(await pending)['@_Id']).toBe('rId1');
		await readPartXml(zip, parser, relsPath);
		await readPartXml(zip, parser, relsPath);
		expect(reads).toHaveBeenCalledTimes(3);
		expect(parses).toHaveBeenCalledTimes(3);
	});
});

describe('load part reader isolation', () => {
	it('never serves one package from the cache of another', async () => {
		const a = new JSZip().file('part.xml', '<a/>');
		const b = new JSZip().file('part.xml', '<b/>');
		const readerA = new LoadPartReader(a);
		const readerB = new LoadPartReader(b);
		expect(await readPartText(a, 'part.xml')).toBe('<a/>');
		expect(await readPartText(b, 'part.xml')).toBe('<b/>');
		readerA.close();
		const readsB = vi.spyOn(b.file('part.xml')!, 'async');
		await readPartText(b, 'part.xml');
		expect(readsB).not.toHaveBeenCalled();
		readerB.close();
	});

	it('leaves packages without an open reader uncached', async () => {
		const zip = new JSZip().file('part.xml', '<a/>');
		const reads = vi.spyOn(zip.file('part.xml')!, 'async');
		await readPartText(zip, 'part.xml');
		await readPartText(zip, 'part.xml');
		expect(reads).toHaveBeenCalledTimes(2);
	});

	it('keeps a newer reader for the same package open when an older one closes', async () => {
		const zip = new JSZip().file('part.xml', '<a/>');
		const older = new LoadPartReader(zip);
		const newer = new LoadPartReader(zip);
		older.close();
		const reads = vi.spyOn(zip.file('part.xml')!, 'async');
		await readPartText(zip, 'part.xml');
		await readPartText(zip, 'part.xml');
		expect(reads).toHaveBeenCalledTimes(1);
		newer.close();
	});
});

describe('load part reader limits', () => {
	async function guarded(maxTotalBytes: number): Promise<JSZip> {
		// Declared sizes exist only once the archive has been written and reopened.
		const bytes = await new JSZip()
			.file('small.xml', '<a/>')
			.file('big.xml', `<a>${'x'.repeat(64)}</a>`)
			.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
		const zip = await JSZip.loadAsync(bytes);
		// A crafted header: big.xml inflates past the size it declares.
		const big = zip.file('big.xml') as unknown as { _data: { uncompressedSize: number } };
		big._data.uncompressedSize = 8;
		return guardZip(zip, { maxEntryBytes: 1024, maxTotalBytes });
	}

	it('reads through the guarded entry and never caches a limit failure', async () => {
		const zip = await guarded(1024);
		const reader = new LoadPartReader(zip);
		await expect(readPartText(zip, 'big.xml')).rejects.toBeInstanceOf(ZipLimitError);
		await expect(readPartText(zip, 'big.xml')).rejects.toBeInstanceOf(ZipLimitError);
		expect(await readPartText(zip, 'small.xml')).toBe('<a/>');
		reader.close();
	});
});
