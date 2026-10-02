import { deflateRawSync } from 'node:zlib';
import JSZip from 'jszip';
import { describe, expect, it, vi } from 'vitest';
import { VisioPackage, VisioPackageError, type VisioPackageLimits } from './package.js';

const encode = (value: string) => new TextEncoder().encode(value);
function crc32(data: Uint8Array): number {
	let crc = 0xffffffff;
	for (const byte of data) {
		crc ^= byte;
		for (let i = 0; i < 8; i++) crc = crc & 1 ? 0xedb88320 ^ (crc >>> 1) : crc >>> 1;
	}
	return (crc ^ 0xffffffff) >>> 0;
}
interface ZipEntry {
	name: string;
	xml?: string;
	data?: Uint8Array;
	deflate?: boolean;
	declaredSize?: number;
	localName?: string;
	flags?: number;
	extra?: Uint8Array;
}
/** Deliberately avoids JSZip's filename sanitization so hostile headers can be tested. */
function archive(specs: ZipEntry[]): Uint8Array {
	const locals: Uint8Array[] = [],
		centrals: Uint8Array[] = [];
	let offset = 0;
	for (const spec of specs) {
		const name = encode(spec.name),
			localName = encode(spec.localName ?? spec.name),
			data = spec.data ?? encode(spec.xml ?? '<Root/>');
		const compressed = spec.deflate ? deflateRawSync(data) : data,
			extra = spec.extra ?? new Uint8Array();
		const local = new Uint8Array(30 + localName.length + extra.length + compressed.length),
			lv = new DataView(local.buffer);
		lv.setUint32(0, 0x04034b50, true);
		lv.setUint16(4, 20, true);
		lv.setUint16(6, spec.flags ?? 0x800, true);
		lv.setUint16(8, spec.deflate ? 8 : 0, true);
		lv.setUint32(14, crc32(data), true);
		lv.setUint32(18, compressed.length, true);
		lv.setUint32(22, spec.declaredSize ?? data.length, true);
		lv.setUint16(26, localName.length, true);
		lv.setUint16(28, extra.length, true);
		local.set(localName, 30);
		local.set(extra, 30 + localName.length);
		local.set(compressed, 30 + localName.length + extra.length);
		const central = new Uint8Array(46 + name.length + extra.length),
			cv = new DataView(central.buffer);
		cv.setUint32(0, 0x02014b50, true);
		cv.setUint16(4, 20, true);
		cv.setUint16(6, 20, true);
		cv.setUint16(8, spec.flags ?? 0x800, true);
		cv.setUint16(10, spec.deflate ? 8 : 0, true);
		cv.setUint32(16, crc32(data), true);
		cv.setUint32(20, compressed.length, true);
		cv.setUint32(24, spec.declaredSize ?? data.length, true);
		cv.setUint16(28, name.length, true);
		cv.setUint16(30, extra.length, true);
		cv.setUint32(42, offset, true);
		central.set(name, 46);
		central.set(extra, 46 + name.length);
		locals.push(local);
		centrals.push(central);
		offset += local.length;
	}
	const centralSize = centrals.reduce((size, bytes) => size + bytes.length, 0),
		end = new Uint8Array(22),
		ev = new DataView(end.buffer);
	ev.setUint32(0, 0x06054b50, true);
	ev.setUint16(8, specs.length, true);
	ev.setUint16(10, specs.length, true);
	ev.setUint32(12, centralSize, true);
	ev.setUint32(16, offset, true);
	const result = new Uint8Array(offset + centralSize + 22);
	let at = 0;
	for (const piece of [...locals, ...centrals, end]) {
		result.set(piece, at);
		at += piece.length;
	}
	return result;
}
const one = (xml = '<Root/>') => archive([{ name: 'visio/document.xml', xml }]);
const openError = (bytes: Uint8Array, code: string, limits?: Partial<VisioPackageLimits>) =>
	expect(VisioPackage.open(bytes, limits)).rejects.toMatchObject({ code });
async function xmlError(xml: string, code = 'INVALID_XML', limits?: Partial<VisioPackageLimits>) {
	const pkg = await VisioPackage.open(one(xml), limits);
	await expect(pkg.readXml('visio/document.xml')).rejects.toMatchObject({ code });
}
const rel = (target: string, mode = '', id = 'r1') =>
	`<Relationship Id="${id}" Type="urn:test" Target="${target}"${mode ? ` TargetMode="${mode}"` : ''}/>`;
const relationships = (body: string) =>
	`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${body}</Relationships>`;

describe('VisioPackage', () => {
	it('exposes validated declared part size without inflating the payload', async () => {
		const bytes = encode('hello');
		const pkg = await VisioPackage.open(
			archive([{ name: 'asset.bin', data: bytes, deflate: true }]),
		);
		const read = vi.spyOn(pkg, 'readBytes');
		expect(pkg.getPartByteLength('asset.bin')).toBe(5);
		expect(read).not.toHaveBeenCalled();
		expect(() => pkg.getPartByteLength('../asset.bin')).toThrow(VisioPackageError);
		expect(() => pkg.getPartByteLength('absent.bin')).toThrow(VisioPackageError);
	});

	it('reads stored or deflated XML without including directory placeholders', async () => {
		for (const streamFiles of [false, true]) {
			const zip = new JSZip();
			zip.file('visio/document.xml', '<?xml version="1.0"?><Root><Text>A &amp; B</Text></Root>');
			const bytes = await zip.generateAsync({
				type: 'uint8array',
				compression: 'DEFLATE',
				streamFiles,
			});
			const pkg = await VisioPackage.open(bytes);
			expect(pkg.paths()).toEqual(['visio/document.xml']);
			expect(pkg.has('visio/')).toBe(false);
			const root = await pkg.readXml('visio/document.xml', 'Root');
			expect(root.textContent).toBe('A & B');
			expect(await pkg.readXml('visio/document.xml')).toBe(root);
		}
	});
	it('accepts ArrayBuffers and owns input bytes after validation', async () => {
		const bytes = one(),
			pkg = await VisioPackage.open(bytes.buffer as ArrayBuffer);
		bytes.fill(0);
		expect((await pkg.readXml('visio/document.xml')).localName).toBe('Root');
	});
	it('decodes UTF-16 XML and rejects invalid UTF-8', async () => {
		const utf16 = Buffer.from('\ufeff<Root>hello</Root>', 'utf16le');
		const pkg = await VisioPackage.open(archive([{ name: 'root.xml', data: utf16 }]));
		expect((await pkg.readXml('root.xml')).textContent).toBe('hello');
		const invalid = await VisioPackage.open(
			archive([{ name: 'root.xml', data: new Uint8Array([0xff, 0xff]) }]),
		);
		await expect(invalid.readXml('root.xml')).rejects.toMatchObject({ code: 'INVALID_XML' });
	});
	it('rejects missing parts and unexpected XML roots', async () => {
		const pkg = await VisioPackage.open(one());
		await expect(pkg.readXml('missing.xml')).rejects.toMatchObject({ code: 'MISSING_PART' });
		await expect(pkg.readXml('visio/document.xml', 'Wrong')).rejects.toMatchObject({
			code: 'INVALID_XML',
		});
		expect(pkg.has('missing.xml')).toBe(false);
	});
	it.each([
		'../x.xml',
		'/x.xml',
		'a/../x.xml',
		'a\\x.xml',
		'a//x.xml',
		'a/./x.xml',
		'%2e%2e/x.xml',
		'a%2fx.xml',
		'a%252fx.xml',
		'http:x.xml',
		'a\u0000.xml',
	])('rejects unsafe entry name %s before JSZip normalization', async (name) => {
		await openError(archive([{ name }]), 'UNSAFE_PATH');
	});
	it('rejects duplicates, URI aliases, and directory/file aliases', async () => {
		for (const names of [
			['x.xml', 'x.xml'],
			['x%20y.xml', 'x y.xml'],
			['x/', 'x'],
		]) {
			await openError(
				archive(names.map((name) => ({ name, xml: name.endsWith('/') ? '' : '<Root/>' }))),
				'DUPLICATE_ENTRY',
			);
		}
	});
	it('rejects local/central name, size, and CRC mismatches', async () => {
		await openError(archive([{ name: 'root.xml', localName: 'evil.xml' }]), 'ZIP_MISMATCH');
		for (const at of [14, 18, 22]) {
			const bytes = one(),
				view = new DataView(bytes.buffer);
			view.setUint32(at, view.getUint32(at, true) + 1, true);
			await openError(bytes, 'ZIP_MISMATCH');
		}
	});
	it('rejects encrypted, ZIP64, multidisk, and unsupported compression entries', async () => {
		await openError(archive([{ name: 'x', flags: 1 }]), 'UNSUPPORTED_ZIP');
		await openError(
			archive([{ name: 'x', extra: new Uint8Array([1, 0, 0, 0]) }]),
			'UNSUPPORTED_ZIP',
		);
		const multidisk = one();
		new DataView(multidisk.buffer).setUint16(multidisk.length - 18, 1, true);
		await openError(multidisk, 'UNSUPPORTED_ZIP');
		const zip64 = one();
		new DataView(zip64.buffer).setUint16(zip64.length - 12, 0xffff, true);
		await openError(zip64, 'UNSUPPORTED_ZIP');
		const method = one(),
			view = new DataView(method.buffer),
			central = view.getUint32(method.length - 6, true);
		view.setUint16(8, 99, true);
		view.setUint16(central + 10, 99, true);
		await openError(method, 'UNSUPPORTED_ZIP');
	});
	it('rejects Unicode path overrides that differ from the validated filename', async () => {
		const path = encode('../evil.xml'),
			extra = new Uint8Array(9 + path.length),
			view = new DataView(extra.buffer);
		view.setUint16(0, 0x7075, true);
		view.setUint16(2, 5 + path.length, true);
		extra[4] = 1;
		extra.set(path, 9);
		await openError(archive([{ name: 'root.xml', extra }]), 'ZIP_MISMATCH');
	});
	it('rejects truncation, missing end records, overlapping entries, and trailing bytes', async () => {
		await openError(new Uint8Array([1, 2, 3]), 'INVALID_ZIP');
		const bytes = one();
		await openError(bytes.slice(0, -1), 'INVALID_ZIP');
		await openError(new Uint8Array([...bytes, 0]), 'INVALID_ZIP');
		const overlap = archive([{ name: 'a.xml' }, { name: 'b.xml' }]),
			view = new DataView(overlap.buffer),
			central = view.getUint32(overlap.length - 6, true);
		view.setUint32(central + 46 + 5 + 42, 0, true);
		await openError(overlap, 'ZIP_MISMATCH');
	});
	it('checks input, count, declared size, aggregate size, and compression ratio limits', async () => {
		await openError(one(), 'LIMIT_INPUT', { maxInputBytes: 1 });
		await openError(archive([{ name: 'a' }, { name: 'b' }]), 'LIMIT_ENTRIES', { maxEntries: 1 });
		await openError(one(), 'LIMIT_ENTRY', { maxEntryBytes: 1 });
		await openError(archive([{ name: 'a' }, { name: 'b' }]), 'LIMIT_TOTAL', { maxTotalBytes: 10 });
		await openError(
			archive([{ name: 'x', xml: `<Root>${'a'.repeat(50000)}</Root>`, deflate: true }]),
			'LIMIT_RATIO',
		);
	});
	it('enforces actual streamed output limits even when declared sizes lie', async () => {
		const bytes = archive([
			{
				name: 'root.xml',
				xml: `<Root>${'a'.repeat(50000)}</Root>`,
				deflate: true,
				declaredSize: 5,
			},
		]);
		const pkg = await VisioPackage.open(bytes, { maxEntryBytes: 100, maxCompressionRatio: 10000 });
		await expect(pkg.readXml('root.xml')).rejects.toMatchObject({ code: 'LIMIT_ENTRY' });
		const ratio = await VisioPackage.open(bytes, { maxCompressionRatio: 10 });
		await expect(ratio.readXml('root.xml')).rejects.toMatchObject({ code: 'LIMIT_RATIO' });
		const total = await VisioPackage.open(bytes, {
			maxTotalBytes: 100,
			maxCompressionRatio: 10000,
		});
		await expect(total.readXml('root.xml')).rejects.toMatchObject({ code: 'LIMIT_TOTAL' });
	});
	it('rejects corrupted payload CRC even when both headers agree', async () => {
		const bytes = one(),
			at = 30 + encode('visio/document.xml').length;
		bytes[at + 1] = 'B'.charCodeAt(0);
		const pkg = await VisioPackage.open(bytes);
		await expect(pkg.readXml('visio/document.xml')).rejects.toMatchObject({ code: 'ZIP_MISMATCH' });
	});
	it.each([
		'<!DOCTYPE Root><Root/>',
		'<!DOCTYPE Root [<!ENTITY x "hello">]><Root>&x;</Root>',
		'<Root>',
		'<Root></Other>',
		'<Root a=unquoted/>',
		'<Root a="1" a="2"/>',
		'<Root/><Other/>',
		'outside<Root/>',
		'<Root>&unknown;</Root>',
		'<Root>&#0;</Root>',
		'<Root><!-- bad -- comment --></Root>',
		'<Root>]]></Root>',
		'<Root><Child/ ></Root>',
	])('rejects malformed or unsafe XML %s', async (xml) => {
		await xmlError(xml);
	});
	it.each([
		'<Root xmlns:x="urn:a" xmlns:y="urn:a" x:z="1" y:z="2"/>',
		'<Root xmlns:xml="urn:wrong"/>',
		'<Root xmlns:other="http://www.w3.org/XML/1998/namespace"/>',
		'<Root xmlns:other="http://www.w3.org/2000/xmlns/"/>',
		'<Root xmlns:other=""/>',
		'<Root\u00a0a="1"/>',
		'\u00a0<Root/>',
		'<Root a="1"b="2"/>',
		'<?xml blah?><Root/>',
		'<?xml version=1.0?><Root/>',
		'<?xml version="1.0"encoding="UTF-8"?><Root/>',
		'<?xml version="1.0" standalone="maybe"?><Root/>',
		'<?xml version="1.2"?><Root/>',
		'<?xml version="1.0" encoding="UTF-7"?><Root/>',
		'<?xml version="1.0" encoding="UTF-16"?><Root/>',
		'<?xml encoding="UTF-8" version="1.0"?><Root/>',
	])('rejects namespace and lexical XML violations %s', async (xml) => {
		await xmlError(xml);
	});
	it('accepts XML declaration whitespace and legitimate reserved namespace use', async () => {
		const pkg = await VisioPackage.open(
			one(
				'<?xml\nversion="1.0"?><Root xmlns:xml="http://www.w3.org/XML/1998/namespace" xml:space="preserve"/>',
			),
		);
		expect((await pkg.readXml('visio/document.xml')).localName).toBe('Root');
	});

	it('accepts namespaces, quoted delimiters, comments, CDATA, and numeric entities', async () => {
		const pkg = await VisioPackage.open(
			one(
				'<v:Root xmlns:v="urn:visio" a="&quot;&gt;">text<!--okay--><![CDATA[<x>&y]]><v:Child/>&#x1f600;</v:Root>',
			),
		);
		expect((await pkg.readXml('visio/document.xml')).textContent).toBe('text<x>&y😀');
	});
	it('checks XML character, node, and depth limits before DOM construction', async () => {
		await xmlError('<Root/>', 'LIMIT_XML', { maxXmlChars: 3 });
		await xmlError('<Root><A/><B/></Root>', 'LIMIT_XML', { maxXmlNodes: 2 });
		await xmlError('<Root><A><B/></A></Root>', 'LIMIT_XML', { maxXmlDepth: 2 });
	});
	it('enforces a processing deadline on open and subsequent reads', async () => {
		const now = Date.now(),
			clock = vi.spyOn(Date, 'now').mockReturnValue(now);
		try {
			const pkg = await VisioPackage.open(one(), { maxRuntimeMs: 100 });
			clock.mockReturnValue(now + 101);
			await expect(pkg.readXml('visio/document.xml')).rejects.toMatchObject({
				code: 'LIMIT_RUNTIME',
			});
			clock.mockImplementationOnce(() => now).mockReturnValue(now + 101);
			await openError(one(), 'LIMIT_RUNTIME', { maxRuntimeMs: 100 });
		} finally {
			clock.mockRestore();
		}
	});
	it('rejects nonpositive or nonfinite limits', async () => {
		for (const maxEntryBytes of [0, -1, NaN, Infinity])
			await openError(one(), 'INVALID_LIMITS', { maxEntryBytes });
		expect(new VisioPackageError('TEST', 'example')).toBeInstanceOf(Error);
	});
	it('resolves root and part relationships while preserving external targets without fetching', async () => {
		const bytes = archive([
			{ name: '_rels/.rels', xml: relationships(rel('visio/document.xml')) },
			{ name: 'visio/document.xml' },
			{ name: 'visio/pages/page1.xml' },
			{ name: 'visio/masters/master1.xml' },
			{ name: 'visio/theme/theme1.xml' },
			{
				name: 'visio/pages/_rels/page1.xml.rels',
				xml: relationships(
					rel('../masters/master1.xml') +
						rel('https://example.com/a?q=1', 'External', 'r2') +
						rel('/visio/theme/theme1.xml', '', 'r3'),
				),
			},
		]);
		const pkg = await VisioPackage.open(bytes);
		expect((await pkg.relationships('')).get('r1')?.target).toBe('visio/document.xml');
		const rels = await pkg.relationships('visio/pages/page1.xml');
		expect(rels.get('r1')?.target).toBe('visio/masters/master1.xml');
		expect(rels.get('r2')).toEqual({
			id: 'r2',
			type: 'urn:test',
			target: 'https://example.com/a?q=1',
			mode: 'External',
		});
		expect(rels.get('r3')?.target).toBe('visio/theme/theme1.xml');
		expect(await pkg.relationships('missing.xml')).toEqual(new Map());
	});
	it('resolves percent-encoded internal targets to validated ZIP names', async () => {
		const pkg = await VisioPackage.open(
			archive([
				{ name: '_rels/.rels', xml: relationships(rel('visio/my%20document.xml')) },
				{ name: 'visio/my%20document.xml' },
			]),
		);
		const target = (await pkg.relationships('')).get('r1')?.target;
		expect(target).toBe('visio/my%20document.xml');
		expect(pkg.has(target ?? '')).toBe(true);
	});

	it.each([
		'../../outside.xml',
		'%2e%2e/%2e%2e/outside.xml',
		'https://example.com/x.xml',
		'file:///etc/passwd',
		'//example.com/x',
		'a\\b.xml',
		'a%252fb.xml',
		'x.xml?query=1',
		'x.xml#fragment',
	])('rejects unsafe internal relationship %s', async (target) => {
		const pkg = await VisioPackage.open(
			archive([{ name: 'visio/_rels/document.xml.rels', xml: relationships(rel(target)) }]),
		);
		await expect(pkg.relationships('visio/document.xml')).rejects.toMatchObject({
			code: 'INVALID_RELATIONSHIP',
		});
	});
	it('rejects missing internal targets and orphan relationship sources', async () => {
		const missing = await VisioPackage.open(
			archive([{ name: '_rels/.rels', xml: relationships(rel('missing.xml')) }]),
		);
		await expect(missing.relationships('')).rejects.toMatchObject({ code: 'INVALID_RELATIONSHIP' });
		const orphan = await VisioPackage.open(
			archive([
				{
					name: 'visio/_rels/document.xml.rels',
					xml: relationships(rel('https://example.com', 'External')),
				},
			]),
		);
		await expect(orphan.relationships('visio/document.xml')).rejects.toMatchObject({
			code: 'INVALID_RELATIONSHIP',
		});
	});

	it('rejects duplicate IDs, invalid modes, and spoofed relationship namespaces', async () => {
		for (const xml of [
			relationships(rel('a') + rel('b')),
			relationships(rel('a', 'external')),
			'<Relationships xmlns="urn:wrong"/>',
			'<Relationships/>',
			relationships('<Relationship Id="r1" Target="x"/>'),
		]) {
			const pkg = await VisioPackage.open(archive([{ name: '_rels/.rels', xml }]));
			await expect(pkg.relationships('')).rejects.toMatchObject({ code: 'INVALID_RELATIONSHIP' });
		}
	});
});

describe('aggregate XML output budget', () => {
	it('bounds total constructed nodes across parts and does not double-charge cached XML', async () => {
		const pkg = await VisioPackage.open(
			archive([{ name: 'a.xml' }, { name: 'b.xml' }, { name: 'c.xml' }, { name: 'd.xml' }]),
			{ maxTotalXmlNodes: 3 },
		);
		await pkg.readXml('a.xml');
		await pkg.readXml('a.xml');
		await pkg.readXml('b.xml');
		await pkg.readXml('c.xml');
		await expect(pkg.readXml('d.xml')).rejects.toMatchObject({ code: 'LIMIT_XML_TOTAL' });
	});
});
