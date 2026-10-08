import { describe, expect, it } from 'vitest';
import {
	captureVisioClipboard,
	deserializeVisioClipboard,
	serializeVisioClipboard,
	VISIO_CLIPBOARD_MAX_CHARS,
} from './clipboard';
import { editVsdx } from './edit';
import { cell, fixture, rectangle, section, shape } from './test-fixtures';
import { snapshotPasteShapes } from './edit-paste-commands';
import type { VisioClipboardSnapshot } from './clipboard-types';
import { assertClipboardXmlLimits } from './clipboard-xml';

const local = (extra = '', attrs = '', x = cell('PinX', 2)) =>
	shape(
		'1',
		x +
			cell('PinY', 3) +
			cell('Width', 2) +
			cell('Height', 1) +
			rectangle +
			'<Text>Clipboard</Text>' +
			extra,
		attrs,
	);
const source = (shapes = local(), document = '') =>
	fixture({ document, pages: [{ id: '0', contents: `<Shapes>${shapes}</Shapes>` }] });
const paste = (bytes: Uint8Array, clipboard: VisioClipboardSnapshot) =>
	editVsdx(bytes, [
		{
			type: 'paste-shapes',
			pageId: '0',
			clipboard,
			copies: clipboard.selectionIds.map((shapeId, index) => ({
				shapeId,
				newShapeId: String(index + 10),
			})),
			offsetX: 0.33,
			offsetY: -0.33,
		},
	]);
describe('clipboard source and hostile payload admission', () => {
	it.each([
		'visio/theme/' + 'x'.repeat(1025) + '.xml',
		'visio/theme/..\\bad.xml',
		'visio/theme/%2e%2e%2fbad.xml',
	])('bounds and validates direct snapshot resource paths', async (path) => {
		const snapshot = await captureVisioClipboard(await source(), '0', ['1']);
		const changed = {
			...snapshot,
			resources: [...snapshot.resources, { path, xml: '<resource/>' }],
		};
		expect(() =>
			snapshotPasteShapes({
				type: 'paste-shapes',
				pageId: '0',
				clipboard: changed,
				copies: [{ shapeId: '1', newShapeId: '2' }],
				offsetX: 0,
				offsetY: 0,
			}),
		).toThrow();
	});
	it('includes resource paths in the direct snapshot aggregate character budget', async () => {
		const snapshot = await captureVisioClipboard(await source(), '0', ['1']);
		const xmlCharacters =
			snapshot.pageContext.length +
			snapshot.resources.reduce((sum, resource) => sum + resource.xml.length, 0);
		const changed = {
			...snapshot,
			shapes: [{ shapeId: '1', xml: 'x'.repeat(VISIO_CLIPBOARD_MAX_CHARS - xmlCharacters) }],
		};
		expect(() =>
			snapshotPasteShapes({
				type: 'paste-shapes',
				pageId: '0',
				clipboard: changed,
				copies: [{ shapeId: '1', newShapeId: '2' }],
				offsetX: 0,
				offsetY: 0,
			}),
		).toThrow();
	});
	it.each(['\uD800', '\uDC00'])(
		'refuses lone Unicode surrogates before serialization or paste',
		async (invalid) => {
			const snapshot = await captureVisioClipboard(await source(), '0', ['1']);
			const changed = {
				...snapshot,
				shapes: [{ shapeId: '1', xml: snapshot.shapes[0]!.xml.replace('Clipboard', invalid) }],
			};
			expect(() => serializeVisioClipboard(changed)).toThrow();
			await expect(paste(await source(), changed)).rejects.toMatchObject({
				code: 'INVALID_CLIPBOARD',
			});
		},
	);
	it('preserves valid Unicode surrogate pairs exactly', async () => {
		const original = await source(local().replace('Clipboard', 'Clipboard🙂'));
		const snapshot = deserializeVisioClipboard(
			serializeVisioClipboard(await captureVisioClipboard(original, '0', ['1'])),
		);
		expect(snapshot.shapes[0]!.xml).toContain('Clipboard🙂');
		expect((await paste(original, snapshot)).changedParts).toHaveLength(1);
	});
	it('refuses nested Visio identities concealed under unknown wrappers', async () => {
		const bytes = await source(local('<Unknown><Shape ID="99"/></Unknown>'));
		await expect(captureVisioClipboard(bytes, '0', ['1'])).rejects.toMatchObject({
			code: 'UNSUPPORTED_DUPLICATE',
		});
		await expect(
			editVsdx(bytes, [
				{
					type: 'duplicate-shapes',
					pageId: '0',
					copies: [{ shapeId: '1', newShapeId: '2' }],
					offsetX: 0,
					offsetY: 0,
				},
			]),
		).rejects.toMatchObject({ code: 'UNSUPPORTED_DUPLICATE' });
	});
	it('bounds aggregate fragments before constructing DOMs or importing any shape', async () => {
		const snapshot = await captureVisioClipboard(await source(), '0', ['1']);
		const xml = `<Shape xmlns="http://schemas.microsoft.com/office/visio/2012/main" ID="1">${'<Unknown/>'.repeat(55_000)}</Shape>`;
		expect(() =>
			assertClipboardXmlLimits(
				{
					...snapshot,
					shapes: [
						{ shapeId: '1', xml },
						{ shapeId: '2', xml: xml.replace('ID="1"', 'ID="2"') },
					],
				},
				() => {},
			),
		).toThrow();
	});
	it.each([
		'<FaceName ID="1" Name="Arial"/><FaceName Name="Calibri"/>',
		'<FaceName ID="0" Name="Arial"/><FaceName ID="1" Name="ARIAL"/>',
	])('refuses ambiguous actual font tables %s', async (faces) => {
		await expect(
			captureVisioClipboard(await source(local(), `<FaceNames>${faces}</FaceNames>`), '0', ['1']),
		).rejects.toMatchObject({ code: 'UNSUPPORTED_CLIPBOARD' });
	});
	it('compares actual theme definitions and theme relationship mapping, not claimed digests', async () => {
		const themed = (value: string, target = 'theme/theme1.xml') =>
			fixture({
				pages: [{ id: '0', contents: `<Shapes>${local()}</Shapes>` }],
				edit: (zip) => {
					zip.file(
						'visio/theme/theme1.xml',
						`<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${value}"/>`,
					);
					zip.file(
						'visio/_rels/document.xml.rels',
						`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/visio/2010/relationships/pages" Target="pages/pages.xml"/><Relationship Id="rId2" Type="http://schemas.microsoft.com/visio/2010/relationships/theme" Target="${target}"/></Relationships>`,
					);
				},
			});
		const original = await themed('Original'),
			snapshot = await captureVisioClipboard(original, '0', ['1']);
		await expect(paste(await themed('Changed'), snapshot)).rejects.toMatchObject({
			code: 'CLIPBOARD_RESOURCE_MISMATCH',
		});
		await expect(
			paste(await themed('Original', './theme/theme1.xml'), snapshot),
		).rejects.toMatchObject({ code: 'CLIPBOARD_RESOURCE_MISMATCH' });
		const forged = {
			...snapshot,
			resources: snapshot.resources.map((resource) => ({ ...resource, sha256: '0'.repeat(64) })),
		};
		expect((await paste(original, forged)).changedParts).toHaveLength(1);
	});
	it.each([
		['master', local('', 'Master="1"')],
		['group', local('<Shapes/>', 'Type="Group"')],
		['foreign', local('<ForeignData/>')],
		['line', local(cell('OneD', 1))],
		['lock', local(cell('LockSelect', 1))],
		['layer', local(cell('LayerMember', '0'))],
		['missing pin', local('', '', '')],
		['error pin', local('', '', '<Cell N="PinX" V="2" E="value"/>')],
		['stale pin', local('', '', cell('PinX', 2, '3'))],
		[
			'dynamic formula',
			local(section('User', '<Row N="A"><Cell N="Value" V="0" F="CONTAINERSHEETREF(1)"/></Row>')),
		],
	] as const)('refuses source %s without changing caller bytes', async (_name, xml) => {
		const bytes = await source(xml),
			before = bytes.slice();
		await expect(captureVisioClipboard(bytes, '0', ['1'])).rejects.toBeInstanceOf(Error);
		expect(bytes).toEqual(before);
	});
	it('admits guarded capture but refuses displaced guarded paste atomically', async () => {
		const bytes = await source(local('', '', cell('PinX', 2, 'GUARD(2)'))),
			before = bytes.slice();
		const snapshot = await captureVisioClipboard(bytes, '0', ['1']);
		await expect(paste(bytes, snapshot)).rejects.toMatchObject({ code: 'EDIT_PROTECTED_CELL' });
		expect(bytes).toEqual(before);
	});
	it('refuses source and tampered dangling inherited fonts against actual resources', async () => {
		const document =
			'<FaceNames><FaceName ID="0" Name="Arial"/></FaceNames><StyleSheets><StyleSheet ID="2"><Section N="Character"><Row IX="0"><Cell N="Font" V="999"/></Row></Section></StyleSheet></StyleSheets>';
		await expect(
			captureVisioClipboard(await source(local('', 'TextStyle="2"'), document), '0', ['1']),
		).rejects.toMatchObject({ code: 'UNSUPPORTED_CLIPBOARD' });
	});
	it('refuses a FONT cache that disagrees with the actual existing family', async () => {
		const document =
			'<FaceNames><FaceName ID="0" Name="Arial"/><FaceName ID="1" Name="Calibri"/></FaceNames>';
		const fonts = section(
			'Character',
			'<Row IX="0"><Cell N="Font" V="0" F="FONT(&quot;Calibri&quot;)"/></Row>',
		);
		await expect(
			captureVisioClipboard(await source(local(fonts), document), '0', ['1']),
		).rejects.toMatchObject({ code: 'UNSUPPORTED_CLIPBOARD' });
	});
	it('preserves quoted pseudo-functions but refuses actual page functions with stale display context', async () => {
		const quoted = section(
			'User',
			'<Row N="A"><Cell N="Value" V="PAGES(" F="&quot;PAGES(&quot;"/></Row>',
		);
		const bytes = await source(local(quoted));
		expect(
			(await paste(bytes, await captureVisioClipboard(bytes, '0', ['1']))).changedParts,
		).toHaveLength(1);
		const pageFunction = section(
			'Field',
			'<Row IX="0"><Cell N="Value" V="1" F="PAGENUMBER()"/></Row>',
		);
		await expect(
			captureVisioClipboard(await source(local(pageFunction)), '0', ['1']),
		).rejects.toMatchObject({ code: 'UNSUPPORTED_CLIPBOARD' });
	});
	it('refuses tampered uncaptured references and pin-dependent cached field display', async () => {
		const bytes = await source(),
			snapshot = await captureVisioClipboard(bytes, '0', ['1']);
		for (const extra of [
			section('User', '<Row N="Missing"><Cell N="Value" V="2" U="DL" F="Sheet.2!PinX"/></Row>'),
			section('Field', '<Row IX="0"><Cell N="Value" V="2" F="PinX/1 in"/></Row>'),
		]) {
			const changed = {
				...snapshot,
				shapes: [
					{ shapeId: '1', xml: snapshot.shapes[0]!.xml.replace('</Shape>', extra + '</Shape>') },
				],
			};
			await expect(paste(bytes, changed)).rejects.toMatchObject({
				code: extra.includes('Field') ? 'UNSUPPORTED_DUPLICATE' : 'UNSUPPORTED_CLIPBOARD',
			});
		}
	});
	it('refuses copied IDs already referenced by target formulas', async () => {
		const initial = await source(),
			snapshot = await captureVisioClipboard(initial, '0', ['1']);
		const target = await source(
			local(section('User', '<Row N="Missing"><Cell N="Value" V="2" F="Sheet.10!PinX"/></Row>')),
		);
		await expect(paste(target, snapshot)).rejects.toMatchObject({
			code: 'EDIT_DUPLICATE_DEPENDENCY',
		});
	});
	it('validates source array/IDs and input bounds before opening a package', async () => {
		for (const ids of [
			[],
			['1', '1'],
			['01'],
			Array.from({ length: 1001 }, (_, index) => String(index)),
		])
			await expect(captureVisioClipboard(new Uint8Array(), '0', ids)).rejects.toBeInstanceOf(Error);
		await expect(
			captureVisioClipboard(new Uint8Array(10), '0', ['1'], { limits: { maxInputBytes: 1 } }),
		).rejects.toMatchObject({ code: 'LIMIT_INPUT' });
	});
	it('strips host fields and requires complete unique paste mappings', async () => {
		const snapshot = await captureVisioClipboard(await source(), '0', ['1']);
		const roundtrip = deserializeVisioClipboard(
			serializeVisioClipboard({ ...snapshot, secret: 'private' } as VisioClipboardSnapshot),
		);
		expect(roundtrip).not.toHaveProperty('secret');
		for (const copies of [
			[],
			[{ shapeId: '2', newShapeId: '10' }],
			[
				{ shapeId: '1', newShapeId: '10' },
				{ shapeId: '1', newShapeId: '11' },
			],
		])
			expect(() =>
				snapshotPasteShapes({
					type: 'paste-shapes',
					pageId: '0',
					clipboard: snapshot,
					copies,
					offsetX: 0,
					offsetY: 0,
				}),
			).toThrow();
	});
});
