import { describe, expect, it, vi } from 'vitest';
import { parseVsdx } from './index.js';
import { fixture, shape, xml, relations } from './test-fixtures.js';
import { emf, record } from './emf-admission.test-fixtures.js';
import { inspectEmbeddedVisioMetafile } from './prepare-metafiles.js';
import { parseXml, NS } from '../xml/index.js';
import type { VisioPackage } from './package.js';

const data = '<ForeignData ForeignType="EnhMetaFile"><Rel r:id="media"/></ForeignData>';
const relation = (mode = 'Internal', type = `${NS.r}/image`) =>
	`<Relationship Id="media" Type="${type}" Target="../media/item.emf" TargetMode="${mode}"/>`;
const build = (bytes: Uint8Array, mode = 'Internal', type = `${NS.r}/image`) =>
	fixture({
		pages: [{ id: '0', contents: `<Shapes>${shape('1', data)}</Shapes>` }],
		edit: (zip) => {
			zip.file('visio/pages/_rels/page1.xml.rels', relations(relation(mode, type)));
			zip.file('visio/media/item.emf', bytes);
		},
	});
const codes = (model: Awaited<ReturnType<typeof parseVsdx>>) =>
	model.diagnostics.map((d) => d.code);
const rectangle = () => emf([record(43, [0, 0, 90, 90])]);

describe('embedded EMF compatibility inspection without conversion', () => {
	it('preserves omission and explicitly reports rendering disabled for admitted structural input', async () => {
		const model = await parseVsdx(await build(rectangle()));
		expect(codes(model)).toContain('emf-rendering-disabled');
		expect(model.pages[0]!.shapes[0]!.image).toBeUndefined();
		expect(model.diagnostics.find((d) => d.code === 'emf-rendering-disabled')).toMatchObject({
			shapeId: '1',
			part: 'visio/media/item.emf',
		});
	});
	it('reports specific unsupported semantics instead of claiming an invalid drawing', async () => {
		const model = await parseVsdx(
			await build(emf([record(38, [1, 6, 3, 0, 0x02000000]), record(43, [0, 0, 90, 90])])),
		);
		expect(codes(model)).toContain('emf-pen-style');
		expect(codes(model)).toContain('emf-palette-color');
		expect(codes(model)).not.toContain('emf-rendering-disabled');
	});
	it('omits malformed EMF while keeping the outer VSDX document usable', async () => {
		const model = await parseVsdx(await build(new Uint8Array([1, 2, 3])));
		expect(codes(model)).toContain('emf-header');
		expect(model.pages).toHaveLength(1);
	});
	it('never loads external image relationships or other relationship categories', async () => {
		const external = await parseVsdx(await build(rectangle(), 'External'));
		expect(codes(external)).toContain('external-image');
		const other = await parseVsdx(await build(rectangle(), 'Internal', `${NS.r}/oleObject`));
		expect(codes(other)).toContain('invalid-image-relationship');
	});
	it('rejects missing and ambiguous relationship references', async () => {
		const model = await parseVsdx(
			await fixture({
				pages: [
					{
						id: '0',
						contents: `<Shapes>${shape('1', data.replace('<Rel r:id="media"/>', '<Rel r:id="media"/><Rel r:id="other"/>'))}</Shapes>`,
					},
				],
			}),
		);
		expect(codes(model)).toContain('invalid-image-relationship');
	});

	it('rejects oversized declared assets before reading or inflating bytes', async () => {
		const readBytes = vi.fn(async () => rectangle());
		const pkg = {
			getPartByteLength: () => 4 * 1024 * 1024 + 1,
			relationships: async () =>
				new Map([
					['media', { mode: 'Internal', target: 'visio/media/big.emf', type: `${NS.r}/image` }],
				]),
			readBytes,
		} as unknown as VisioPackage;
		const element = parseXml(xml('Shape', data)).documentElement.firstChild as Element;
		const notes: string[] = [];
		await inspectEmbeddedVisioMetafile(pkg, 'page', element, (code) => notes.push(code));
		expect(readBytes).not.toHaveBeenCalled();
		expect(notes).toEqual(['emf-input-limit']);
	});
	it('deduplicates scanning of a shared internal media part across instances', async () => {
		const readBytes = vi.fn(async () => rectangle());
		const pkg = {
			getPartByteLength: () => rectangle().byteLength,
			relationships: async () =>
				new Map([
					['media', { mode: 'Internal', target: 'visio/media/one.emf', type: `${NS.r}/image` }],
				]),
			readBytes,
		} as unknown as VisioPackage;
		const element = parseXml(xml('Shape', data)).documentElement.firstChild as Element;
		const notes: string[] = [];
		await inspectEmbeddedVisioMetafile(pkg, 'page', element, (code) => notes.push(code));
		await inspectEmbeddedVisioMetafile(pkg, 'page', element, (code) => notes.push(code));
		expect(readBytes).toHaveBeenCalledTimes(1);
		expect(notes).toEqual(['emf-rendering-disabled', 'emf-rendering-disabled']);
	});

	it('reserves concurrent slots and coalesces pending reads without bypassing the asset limit', async () => {
		let next = 0;
		const readBytes = vi.fn(async () => {
			await Promise.resolve();
			return rectangle();
		});
		const pkg = {
			getPartByteLength: () => rectangle().byteLength,
			relationships: async () =>
				new Map([
					[
						'media',
						{
							mode: 'Internal',
							target: `visio/media/${Math.floor(next++ / 2)}.emf`,
							type: `${NS.r}/image`,
						},
					],
				]),
			readBytes,
		} as unknown as VisioPackage;
		const element = parseXml(xml('Shape', data)).documentElement.firstChild as Element;
		const notes: string[] = [];
		await Promise.all(
			Array.from({ length: 80 }, () =>
				inspectEmbeddedVisioMetafile(pkg, 'page', element, (code) => notes.push(code)),
			),
		);
		expect(readBytes).toHaveBeenCalledTimes(32);
		expect(notes.filter((code) => code === 'emf-document-limit')).toHaveLength(16);
	});
	it('stops materialization before crossing aggregate byte and record limits', async () => {
		const element = parseXml(xml('Shape', data)).documentElement.firstChild as Element;
		let next = 0;
		const large = new Uint8Array(4 * 1024 * 1024);
		const readBytes = vi.fn(async () => large);
		const pkg = {
			getPartByteLength: () => large.length,
			relationships: async () =>
				new Map([
					[
						'media',
						{ mode: 'Internal', target: `visio/media/${next++}.emf`, type: `${NS.r}/image` },
					],
				]),
			readBytes,
		} as unknown as VisioPackage;
		const notes: string[] = [];
		await Promise.all(
			Array.from({ length: 12 }, () =>
				inspectEmbeddedVisioMetafile(pkg, 'page', element, (code) => notes.push(code)),
			),
		);
		expect(readBytes).toHaveBeenCalledTimes(8);
		expect(notes.filter((code) => code === 'emf-document-limit')).toHaveLength(4);
		let id = 0;
		const many = emf(Array.from({ length: 19998 }, () => record(18, [1])));
		const readRecords = vi.fn(async () => many);
		const second = {
			getPartByteLength: () => many.length,
			relationships: async () =>
				new Map([
					['media', { mode: 'Internal', target: `visio/media/${id++}.emf`, type: `${NS.r}/image` }],
				]),
			readBytes: readRecords,
		} as unknown as VisioPackage;
		await Promise.all(
			Array.from({ length: 9 }, () =>
				inspectEmbeddedVisioMetafile(second, 'page', element, () => {}),
			),
		);
		expect(readRecords).toHaveBeenCalledTimes(5);
	});
	it('bounds unique asset reads before materializing later parts', async () => {
		let next = 0;
		const readBytes = vi.fn(async () => rectangle());
		const pkg = {
			getPartByteLength: () => rectangle().byteLength,
			relationships: async () =>
				new Map([
					[
						'media',
						{ mode: 'Internal', target: `visio/media/${next++}.emf`, type: `${NS.r}/image` },
					],
				]),
			readBytes,
		} as unknown as VisioPackage;
		const element = parseXml(xml('Shape', data)).documentElement.firstChild as Element;
		const notes: string[] = [];
		for (let i = 0; i < 40; i++)
			await inspectEmbeddedVisioMetafile(pkg, 'page', element, (code) => notes.push(code));
		expect(readBytes).toHaveBeenCalledTimes(32);
		expect(notes.filter((code) => code === 'emf-document-limit')).toHaveLength(8);
	});
});
