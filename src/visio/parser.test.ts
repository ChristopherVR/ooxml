import { describe, expect, it } from 'vitest';
import { getVisioPageLayers, parseVsdx } from './index.js';
import {
	cell,
	fixture,
	rectangle,
	relation,
	relations,
	row,
	section,
	shape,
} from './test-fixtures.js';

const pageShapes = (contents: string) => `<Shapes>${contents}</Shapes>`;
describe('VSDX scene parsing', () => {
	it('parses page dimensions, rectangle geometry, inert text and cached styles', async () => {
		const bytes = await fixture({
			pages: [
				{
					id: '0',
					width: 12,
					height: 6,
					contents: pageShapes(
						shape(
							'1',
							cell('Width', 4) +
								cell('Height', 2) +
								cell('PinX', 6) +
								cell('PinY', 3) +
								cell('FillForegnd', '#ff5500') +
								cell('LineWeight', 0.025) +
								section('Character', row(0, '', cell('Size', 0.25) + cell('Style', 3))) +
								rectangle +
								'<Text>Hello &lt;script&gt; &amp; world</Text>',
						),
					),
				},
			],
		});
		const document = await parseVsdx(bytes),
			page = document.pages[0]!,
			item = page.shapes[0]!;
		expect(page).toMatchObject({ id: '0', width: 12, height: 6, isBackground: false });
		expect(item.transform).toEqual([1, 0, -0, 1, 4, 2]);
		expect(item.geometry[0]).toEqual({
			path: 'M 0 0 L 4 0 L 4 2 L 0 2 L 0 0',
			fill: true,
			stroke: true,
		});
		expect(item.text.plainText).toBe('Hello <script> & world');
		expect(item.text.runs[0]).toMatchObject({ bold: true, italic: true, fontSize: 0.25 });
		expect(item.style).toMatchObject({ fill: '#ff5500', lineWidth: 0.025 });
	});
	it('keeps groups local and applies rotation, flips and explicit local pins', async () => {
		const nested = shape(
			'2',
			cell('Width', 2) + cell('Height', 1) + cell('PinX', 1) + cell('PinY', 1) + rectangle,
		);
		const document = await parseVsdx(
			await fixture({
				pages: [
					{
						id: '0',
						contents: pageShapes(
							shape(
								'1',
								cell('Width', 4) +
									cell('Height', 3) +
									cell('PinX', 5) +
									cell('PinY', 6) +
									cell('LocPinX', 0) +
									cell('LocPinY', 0) +
									cell('Angle', Math.PI / 2) +
									cell('FlipX', 1) +
									`<Shapes>${nested}</Shapes>`,
								'Type="Group"',
							),
						),
					},
				],
			}),
		);
		const group = document.pages[0]!.shapes[0]!;
		expect(group.kind).toBe('group');
		expect(group.transform[1]).toBeCloseTo(-1);
		expect(group.transform[2]).toBeCloseTo(-1);
		expect(group.transform.slice(4)).toEqual([5, 6]);
		expect(group.children[0]!.transform.slice(4)).toEqual([0, 0.5]);
	});
	it('merges master cells, geometry row overrides, text and deleted rows', async () => {
		const master = shape(
			'10',
			cell('Width', 2) +
				cell('Height', 1) +
				cell('FillForegnd', '#004488') +
				rectangle +
				'<Text>Inherited</Text>',
		);
		const local = shape(
			'3',
			cell('Width', 5) +
				section('Geometry', '<Row IX="3"><Cell N="Y" V="0.5"/></Row><Row IX="4" Del="1"/>'),
			'Master="7"',
		);
		const document = await parseVsdx(
			await fixture({
				masters: [{ id: '7', shapes: master }],
				pages: [{ id: '0', contents: pageShapes(local) }],
			}),
		);
		const item = document.pages[0]!.shapes[0]!;
		expect(item.geometry[0]!.path).toBe('M 0 0 L 5 0 L 5 0.5 L 0 0');
		expect(item.text.plainText).toBe('Inherited');
		expect(item.style.fill).toBe('#004488');
		expect(item.masterId).toBe('7');
	});
	it('inherits master subshapes and overrides the matching MasterShape only', async () => {
		const master = shape('10', `<Shapes>${shape('11')}${shape('12')}</Shapes>`, 'Type="Group"');
		const local = shape(
			'20',
			`<Shapes>${shape('21', '<Text>Changed</Text>', 'MasterShape="11"')}</Shapes>`,
			'Master="5"',
		);
		const document = await parseVsdx(
			await fixture({
				masters: [{ id: '5', shapes: master }],
				pages: [{ id: '0', contents: pageShapes(local) }],
			}),
		);
		const group = document.pages[0]!.shapes[0]!;
		expect(group.children.map((c) => c.id)).toEqual(['21', '20:master:12']);
		expect(group.children[0]!.text.plainText).toBe('Changed');
		expect(group.children[0]!.geometry).toHaveLength(1);
	});
	it('resolves style chains, colors, font table and text runs', async () => {
		const documentXml = `<Colors><ColorEntry IX="20" RGB="#aabbcc"/></Colors><FaceNames><FaceName ID="9" Name="Example Sans"/></FaceNames><StyleSheets><StyleSheet ID="1">${cell('FillForegnd', 20)}${section('Character', row(0, '', cell('Font', 9) + cell('Size', 0.2)) + row(1, '', cell('Style', 1) + cell('Size', 0.3)))}</StyleSheet><StyleSheet ID="2" FillStyle="1"/></StyleSheets>`;
		const result = await parseVsdx(
			await fixture({
				document: documentXml,
				pages: [
					{
						id: '0',
						contents: pageShapes(
							shape('1', rectangle + '<Text>A<cp IX="1"/>B</Text>', 'FillStyle="2" TextStyle="1"'),
						),
					},
				],
			}),
		);
		const item = result.pages[0]!.shapes[0]!;
		expect(item.style.fill).toBe('#aabbcc');
		expect(item.text.fontFamily).toBe('Example Sans');
		expect(item.text.runs.map((run) => [run.text, run.bold])).toEqual([
			['A', false],
			['B', true],
		]);
	});
	it('preserves connector attachment metadata', async () => {
		const contents =
			pageShapes(shape('1') + shape('2', cell('OneD', 1) + rectangle)) +
			'<Connects><Connect FromSheet="2" FromCell="BeginX" FromPart="9" ToSheet="1" ToCell="PinX" ToPart="3"/></Connects>';
		const document = await parseVsdx(await fixture({ pages: [{ id: '0', contents }] }));
		expect(document.pages[0]!.shapes[1]!.kind).toBe('connector');
		expect(document.pages[0]!.connectors[0]).toEqual({
			fromShapeId: '2',
			toShapeId: '1',
			fromCell: 'BeginX',
			toCell: 'PinX',
			fromPart: 9,
			toPart: 3,
		});
	});
	it('returns background layers in paint order and diagnoses cycles', async () => {
		const document = await parseVsdx(
			await fixture({
				pages: [
					{ id: '1', contents: '', attributes: 'BackPage="2"' },
					{ id: '2', contents: '', attributes: 'Background="1" BackPage="3"' },
					{ id: '3', contents: '', attributes: 'Background="1"' },
				],
			}),
		);
		expect(getVisioPageLayers(document, '1').map((page) => page.id)).toEqual(['3', '2', '1']);
		const cyclic = await parseVsdx(
			await fixture({
				pages: [
					{ id: '1', contents: '', attributes: 'BackPage="2"' },
					{ id: '2', contents: '', attributes: 'BackPage="1"' },
				],
			}),
		);
		expect(cyclic.diagnostics.some((d) => d.code === 'background-cycle')).toBe(true);
		expect(getVisioPageLayers(cyclic, '1')).toHaveLength(1);
	});
	it('uses cached formula values without executing formulas and diagnoses missing cache', async () => {
		const data = shape(
			'1',
			cell('Width', 3, 'RUNADDON(&quot;evil&quot;)') + '<Cell N="Height" F="Width*2"/>' + rectangle,
		);
		const document = await parseVsdx(
			await fixture({ pages: [{ id: '0', contents: pageShapes(data) }] }),
		);
		expect(document.pages[0]!.shapes[0]!.width).toBe(3);
		expect(document.diagnostics.some((d) => d.code === 'missing-cached-value')).toBe(true);
	});
	it('reports unsupported geometry, media, themed colors and fills', async () => {
		const data = shape(
			'1',
			cell('FillForegnd', 'Themed') +
				cell('FillPattern', 25) +
				section('Geometry', row(1, 'NURBSTo', '')) +
				'<ForeignData ForeignType="Bitmap"/>',
		);
		const document = await parseVsdx(
			await fixture({ pages: [{ id: '0', contents: pageShapes(data) }] }),
		);
		expect(document.pages[0]!.shapes[0]!.geometry).toEqual([]);
		expect(document.diagnostics.map((d) => d.code)).toEqual(
			expect.arrayContaining([
				'invalid-geometry',
				'unsupported-foreign-object',
				'unsupported-color',
				'unsupported-fill-pattern',
			]),
		);
	});
	it('rejects wrong namespaces, roots and other Office packages', async () => {
		await expect(
			parseVsdx(
				await fixture({
					edit: (zip) => zip.file('visio/pages/page1.xml', '<PageContents xmlns="urn:evil"/>'),
				}),
			),
		).rejects.toMatchObject({ code: 'INVALID_VISIO_XML' });
		await expect(
			parseVsdx(
				await fixture({
					edit: (zip) =>
						zip.file(
							'[Content_Types].xml',
							'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
						),
				}),
			),
		).rejects.toMatchObject({ code: 'UNSUPPORTED_FORMAT' });
	});
	it('rejects absent, mismatched and external structural relationships', async () => {
		for (const value of [
			'',
			relation('rId1', 'master', 'page1.xml'),
			relation('rId1', 'page', 'https://example.com/page.xml', 'External'),
		])
			await expect(
				parseVsdx(
					await fixture({
						edit: (zip) => zip.file('visio/pages/_rels/pages.xml.rels', relations(value)),
					}),
				),
			).rejects.toMatchObject({ code: 'INVALID_RELATIONSHIP' });
	});
	it('rejects duplicate page/shape IDs and invalid dimensions', async () => {
		await expect(
			parseVsdx(
				await fixture({
					pages: [
						{ id: '0', contents: '' },
						{ id: '0', contents: '' },
					],
				}),
			),
		).rejects.toMatchObject({ code: 'INVALID_PAGE_ID' });
		await expect(
			parseVsdx(
				await fixture({ pages: [{ id: '0', contents: pageShapes(shape('1') + shape('1')) }] }),
			),
		).rejects.toMatchObject({ code: 'INVALID_SHAPE_ID' });
		await expect(
			parseVsdx(await fixture({ pages: [{ id: '0', contents: '', width: 0 }] })),
		).rejects.toMatchObject({ code: 'INVALID_PAGE_SIZE' });
	});
	it('bounds shape expansion, geometry, text, nesting, and diagnostics', async () => {
		await expect(
			parseVsdx(
				await fixture({ pages: [{ id: '0', contents: pageShapes(shape('1') + shape('2')) }] }),
				{ maxShapes: 1 },
			),
		).rejects.toMatchObject({ code: 'SHAPE_LIMIT' });
		await expect(parseVsdx(await fixture(), { maxGeometryCommands: 1 })).rejects.toMatchObject({
			code: 'GEOMETRY_LIMIT',
		});
		await expect(
			parseVsdx(
				await fixture({
					pages: [{ id: '0', contents: pageShapes(shape('1', '<Text>Too much</Text>')) }],
				}),
				{ maxTextCharacters: 2 },
			),
		).rejects.toMatchObject({ code: 'TEXT_LIMIT' });
		const nested = shape('1', `<Shapes>${shape('2', `<Shapes>${shape('3')}</Shapes>`)}</Shapes>`);
		await expect(
			parseVsdx(await fixture({ pages: [{ id: '0', contents: pageShapes(nested) }] }), {
				maxShapeDepth: 1,
			}),
		).rejects.toMatchObject({ code: 'SHAPE_DEPTH_LIMIT' });
		const capped = await parseVsdx(
			await fixture({ pages: [{ id: '0', contents: pageShapes(shape('1', '<ForeignData/>')) }] }),
			{ maxDiagnostics: 1 },
		);
		expect(capped.diagnostics).toHaveLength(1);
		expect(capped.diagnostics[0]!.code).toBe('diagnostics-truncated');
	});
	it('derives an un-routed connector from cached endpoints when XForm is absent', async () => {
		const line = shape(
			'1',
			cell('BeginX', 2) + cell('BeginY', 3) + cell('EndX', 5) + cell('EndY', 7),
		);
		const document = await parseVsdx(
			await fixture({ pages: [{ id: '0', contents: pageShapes(line) }] }),
		);
		const item = document.pages[0]!.shapes[0]!;
		expect(item.width).toBe(5);
		expect(item.height).toBe(0);
		expect(item.geometry[0]!.path).toBe('M 0 0 L 5 0');
		expect(item.transform[4]).toBeCloseTo(2);
		expect(item.transform[5]).toBeCloseTo(3);
		expect(item.transform[0]).toBeCloseTo(0.6);
		expect(item.transform[1]).toBeCloseTo(0.8);
	});
	it('bounds inherited expansion before allocating an oversized scene', async () => {
		const master = shape('10', `<Shapes>${shape('11')}${shape('12')}${shape('13')}</Shapes>`);
		const bytes = await fixture({
			masters: [{ id: '5', shapes: master }],
			pages: [{ id: '0', contents: pageShapes(shape('1', '', 'Master="5"')) }],
		});
		await expect(parseVsdx(bytes, { maxShapes: 2 })).rejects.toMatchObject({ code: 'SHAPE_LIMIT' });
	});
	it('matches local children against multiple top-level master shapes', async () => {
		const master = shape('10') + shape('11');
		const local = shape(
			'1',
			`<Shapes>${shape('2', '<Text>Override</Text>', 'MasterShape="11"')}</Shapes>`,
			'Master="5"',
		);
		const document = await parseVsdx(
			await fixture({
				masters: [{ id: '5', shapes: master }],
				pages: [{ id: '0', contents: pageShapes(local) }],
			}),
		);
		const group = document.pages[0]!.shapes[0]!;
		expect(group.children.map((shape) => shape.id)).toEqual(['1:master:10', '2']);
		expect(group.children[1]!.text.plainText).toBe('Override');
		expect(group.children[1]!.geometry).toHaveLength(1);
	});
	it('honors the cached HideText cell', async () => {
		const document = await parseVsdx(
			await fixture({
				pages: [
					{
						id: '0',
						contents: pageShapes(shape('1', cell('HideText', 1) + '<Text>Hidden</Text>')),
					},
				],
			}),
		);
		expect(document.pages[0]!.shapes[0]!.text.plainText).toBe('');
	});
	it('ignores MasterShape when a child has its own Master attribute', async () => {
		const masters = [
			{
				id: '5',
				shapes: shape(
					'10',
					`<Shapes>${shape('11', cell('FillForegnd', '#ff0000'))}</Shapes>`,
					'Type="Group"',
				),
			},
			{ id: '6', shapes: shape('20', cell('FillForegnd', '#00ff00')) },
		];
		const local = shape(
			'1',
			`<Shapes>${shape('2', '', 'MasterShape="11" Master="6"')}</Shapes>`,
			'Master="5"',
		);
		const document = await parseVsdx(
			await fixture({ masters, pages: [{ id: '0', contents: pageShapes(local) }] }),
		);
		const children = document.pages[0]!.shapes[0]!.children;
		expect(children).toHaveLength(2);
		expect(children[0]!.style.fill).toBe('#ff0000');
		expect(children[1]!.style.fill).toBe('#00ff00');
	});
	it.each([0, 1, 2] as const)(
		'preserves cached group DisplayMode %s without hiding its children',
		async (mode) => {
			const data = shape(
				'1',
				cell('DisplayMode', mode) + `<Text>Group</Text><Shapes>${shape('2')}</Shapes>`,
				'Type="Group"',
			);
			const document = await parseVsdx(
				await fixture({ pages: [{ id: '0', contents: pageShapes(data) }] }),
			);
			const group = document.pages[0]!.shapes[0]!;
			expect(group.groupDisplayMode).toBe(mode);
			expect(group.hidden).toBe(false);
			expect(group.children[0]!.hidden).toBe(false);
		},
	);
	it('keeps absent group display mode unspecified and inherits an explicit master mode', async () => {
		const master = shape(
			'10',
			cell('DisplayMode', 2) + `<Shapes>${shape('11')}</Shapes>`,
			'Type="Group"',
		);
		const data =
			shape('1', `<Shapes>${shape('2')}</Shapes>`, 'Type="Group"') + shape('3', '', 'Master="5"');
		const document = await parseVsdx(
			await fixture({
				masters: [{ id: '5', shapes: master }],
				pages: [{ id: '0', contents: pageShapes(data) }],
			}),
		);
		expect(document.pages[0]!.shapes[0]!.groupDisplayMode).toBeUndefined();
		expect(document.pages[0]!.shapes[1]!.groupDisplayMode).toBe(2);
	});
	it('validates otherwise unused root-level relationship parts', async () => {
		const bytes = await fixture({
			edit: (zip) => {
				zip.file('unused.xml', '<Root/>');
				zip.file('_rels/unused.xml.rels', relations(relation('rId1', 'image', '../outside.xml')));
			},
		});
		await expect(parseVsdx(bytes)).rejects.toMatchObject({ code: 'INVALID_RELATIONSHIP' });
	});
});
