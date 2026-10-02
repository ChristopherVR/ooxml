import { describe, expect, it } from 'vitest';
import { parseXml } from '../xml/index.js';
import { parseVsdx } from './index.js';
import { indexLayers, pageLayers, shapeLayers } from './layers.js';
import { readSheet } from './sheet.js';
import { cell, fixture, row, section, shape, xml } from './test-fixtures.js';

const resources = { colors: new Map<string, string>(), fonts: new Map<string, string>() };
const sheet = (data: string) => readSheet(parseXml(xml('PageSheet', data)).documentElement);
describe('Visio cached layers', () => {
	it('reads indexed page layers and preserves color, visibility, print and lock metadata', () => {
		const layers = pageLayers(
			sheet(
				section(
					'Layer',
					row(
						0,
						'',
						cell('Name', 'Engineering') +
							cell('Visible', 0) +
							cell('Print', 1) +
							cell('Lock', 1) +
							cell('Color', '#ff0000') +
							cell('ColorTrans', 0.4),
					),
				),
			),
			resources,
			() => {},
		);
		expect(layers).toEqual([
			{
				id: '0',
				name: 'Engineering',
				visible: false,
				printable: true,
				locked: true,
				color: '#ff0000',
				colorOpacity: 0.6,
			},
		]);
	});
	it('honors no-color sentinel255 and default layer attributes', () => {
		const layers = pageLayers(
			sheet(section('Layer', row(2, '', cell('NameUniv', 'Layer 3') + cell('Color', 255)))),
			resources,
			() => {},
		);
		expect(layers[0]).toEqual({
			id: '2',
			name: 'Layer 3',
			visible: true,
			printable: true,
			locked: false,
		});
	});
	it('hides membership on any invisible layer according to MS-VSDX 2.2.3.2.2', () => {
		const layers = pageLayers(
			sheet(section('Layer', row(0, '', cell('Visible', 0)) + row(1, '', cell('Visible', 1)))),
			resources,
			() => {},
		);
		expect(shapeLayers(sheet(cell('LayerMember', '0;1')), indexLayers(layers), () => {})).toEqual({
			layerIds: ['0', '1'],
			hidden: true,
			printSummary: 'all-enabled',
		});
		expect(shapeLayers(sheet(cell('LayerMember', '1')), indexLayers(layers), () => {}).hidden).toBe(
			false,
		);
	});
	it('reports missing or malformed memberships and unapplied colors', () => {
		const messages: string[] = [],
			report = (code: string) => messages.push(code);
		const layers = pageLayers(
			sheet(section('Layer', row(0, '', cell('Color', '#ff0000')))),
			resources,
			report,
		);
		expect(
			shapeLayers(sheet(cell('LayerMember', 'RUN()')), indexLayers(layers), report).layerIds,
		).toEqual([]);
		shapeLayers(sheet(cell('LayerMember', '0;99')), indexLayers(layers), report);
		expect(messages).toEqual([
			'invalid-layer-membership',
			'missing-layer',
			'unsupported-layer-color',
		]);
	});
	it('integrates page layer metadata and inherited shape membership', async () => {
		const bytes = await fixture({
			masters: [{ id: '5', shapes: shape('10', cell('LayerMember', '0')) }],
			pages: [{ id: '1', contents: `<Shapes>${shape('2', '', 'Master="5"')}</Shapes>` }],
			edit: (zip) => {
				zip.file(
					'visio/pages/pages.xml',
					xml(
						'Pages',
						`<Page ID="1"><PageSheet>${cell('PageWidth', 8)}${cell('PageHeight', 6)}${section('Layer', row(0, '', cell('Name', 'Hidden') + cell('Visible', 0)))}</PageSheet><Rel r:id="rId1"/></Page>`,
					),
				);
			},
		});
		const document = await parseVsdx(bytes);
		expect(document.pages[0]!.layers![0]!.name).toBe('Hidden');
		expect(document.pages[0]!.shapes[0]!.layerIds).toEqual(['0']);
		expect(document.pages[0]!.shapes[0]!.hidden).toBe(true);
	});
});
