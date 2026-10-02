import { describe, expect, it } from 'vitest';
import { parseVsdx, resolveVisioPageVisibility } from './index.js';
import { cell, fixture, rectangle, row, section, shape, xml } from './test-fixtures.js';

async function scene(contents: string, layers: string, masters?: { id: string; shapes: string }[]) {
	return parseVsdx(
		await fixture({
			pages: [{ id: '1', contents: `<Shapes>${contents}</Shapes>` }],
			...(masters ? { masters } : {}),
			edit(zip) {
				zip.file(
					'visio/pages/pages.xml',
					xml(
						'Pages',
						`<Page ID="1"><PageSheet>${section('Layer', layers)}</PageSheet><Rel r:id="rId1"/></Page>`,
					),
				);
			},
		}),
	);
}
const layers = (visible: number, printable: number) =>
	row(0, '', cell('Visible', visible) + cell('Print', printable));

describe('Visio independent saved visibility metadata', () => {
	const combinations = [0, 1].flatMap((visible) =>
		[0, 1].flatMap((printable) =>
			[0, 1].flatMap((guide) => [0, 1].map((noShow) => ({ visible, printable, guide, noShow }))),
		),
	);
	it.each(combinations)(
		'preserves display reasons independently for %j',
		async ({ visible, printable, guide, noShow }) => {
			const doc = await scene(
				shape(
					'1',
					cell('LayerMember', '0') + cell('NoShow', noShow) + rectangle,
					guide ? 'Type="Guide"' : '',
				),
				layers(visible, printable),
			);
			const item = doc.pages[0]!.shapes[0]!;
			expect(item.hidden).toBe(!visible || !!guide || !!noShow);
			expect(item.visibility).toEqual({
				layerHidden: !visible,
				guide: !!guide,
				noShow: !!noShow,
				layerPrintSummary: printable ? 'all-enabled' : 'all-disabled',
			});
			expect(resolveVisioPageVisibility(doc.pages[0]!)[0]!.hidden).toBe(item.hidden);
		},
	);
	it.each([
		['', 'unlayered'],
		[cell('LayerMember', ''), 'unlayered'],
		[cell('LayerMember', ' '), 'unlayered'],
		[cell('LayerMember', '0'), 'all-enabled'],
		[cell('LayerMember', '1'), 'all-disabled'],
		[cell('LayerMember', '0;00;2'), 'all-enabled'],
		[cell('LayerMember', '1;3'), 'all-disabled'],
		[cell('LayerMember', '0;1'), 'mixed'],
		[cell('LayerMember', '0;1;99'), 'unknown'],
		[cell('LayerMember', '99'), 'unknown'],
		[cell('LayerMember', 'RUN()'), 'unknown'],
		['<Cell N="LayerMember" F="RUN()"/>', 'unknown'],
	])(
		'summarizes membership facts for %s without deciding printability',
		async (membership, expected) => {
			const doc = await scene(
				shape('1', membership + rectangle),
				[1, 0, 1, 0].map((printable, id) => row(id, '', cell('Print', printable))).join(''),
			);
			expect(doc.pages[0]!.shapes[0]!.visibility!.layerPrintSummary).toBe(expected);
		},
	);
	it.each([
		[cell('NonPrinting', 0), false],
		[cell('NonPrinting', 1), true],
		[cell('NonPrinting', -1), true],
		['', undefined],
		[cell('NonPrinting', 'broken'), undefined],
		['<Cell N="NonPrinting" F="RUN()"/>', undefined],
	])('retains only independently usable NonPrinting caches: %s', async (printing, expected) => {
		const doc = await scene(shape('1', printing + rectangle, 'Type="Guide"'), '');
		const item = doc.pages[0]!.shapes[0]!;
		expect(item.hidden).toBe(true);
		expect(item.visibility!.guide).toBe(true);
		expect(item.visibility!.nonPrinting).toBe(expected);
		expect(Object.hasOwn(item.visibility!, 'nonPrinting')).toBe(expected !== undefined);
	});
	it('keeps geometry NoShow independent of shape suppression and text', async () => {
		const doc = await scene(
			shape(
				'1',
				rectangle.replace(
					'<Section N="Geometry" IX="0">',
					'<Section N="Geometry" IX="0">' + cell('NoShow', 1),
				) + '<Text>Still visible text</Text>',
			),
			'',
		);
		const item = doc.pages[0]!.shapes[0]!;
		expect(item.hidden).toBe(false);
		expect(item.visibility!.noShow).toBe(false);
		expect(item.geometry).toEqual([]);
		expect(item.text.plainText).toBe('Still visible text');
	});
	it('does not reveal a layer-hidden shape with an unusable top-level NoShow cache', async () => {
		const doc = await scene(
			shape('1', cell('LayerMember', '0') + '<Cell N="NoShow" F="RUN()"/>' + rectangle),
			layers(0, 1),
		);
		const page = doc.pages[0]!;
		expect(page.shapes[0]!.visibility!.noShow).toBeUndefined();
		expect(page.shapes[0]!.hidden).toBe(true);
		expect(
			resolveVisioPageVisibility(page, { layerVisibilityOverrides: new Map([['0', true]]) })[0]!
				.hidden,
		).toBe(true);
		// The old short-circuit did not report an unusable NoShow on already-hidden shapes.
		expect(doc.diagnostics.some((entry) => entry.message.includes('Cell NoShow'))).toBe(false);
	});
	it('uses cached master inheritance and local overrides without inheriting group membership into children', async () => {
		const doc = await scene(
			shape(
				'2',
				`<Shapes>${shape('3', cell('NoShow', 0), 'MasterShape="11"')}</Shapes>`,
				'Master="5"',
			),
			layers(0, 1),
			[
				{
					id: '5',
					shapes: shape(
						'10',
						cell('LayerMember', '0') +
							'<Shapes>' +
							shape('11', cell('NoShow', 1) + cell('NonPrinting', 1) + rectangle) +
							'</Shapes>',
						'Type="Group"',
					),
				},
			],
		);
		const parent = doc.pages[0]!.shapes[0]!;
		const child = parent.children[0]!;
		expect(parent.visibility!.layerHidden).toBe(true);
		expect(parent.hidden).toBe(true);
		expect(child.layerIds).toEqual([]);
		expect(child.visibility).toMatchObject({
			layerHidden: false,
			noShow: false,
			nonPrinting: true,
		});
		expect(child.hidden).toBe(false);
		expect(resolveVisioPageVisibility(doc.pages[0]!).map((item) => item.hidden)).toEqual([
			true,
			true,
		]);
		expect(
			resolveVisioPageVisibility(doc.pages[0]!, {
				layerVisibilityOverrides: new Map([['0', true]]),
			}).map((item) => item.hidden),
		).toEqual([false, false]);
	});
});
