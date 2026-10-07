import type { DiagramDrawing, DiagramDrawingShape } from 'ooxml-core/diagram';
import { readFileSync } from 'node:fs';
import { URL as NodeURL } from 'node:url';
import { loadXlsx } from 'ooxml-core/xlsx';
import { createTestContext } from '../xlsx/commands/test-support';
import { paintSmartArt } from '../xlsx/grid/smartart';
import { colorToCss } from './smartart-svg';
import { registerOfficeUi, type SmartArtRenderReport } from '../index';

beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

const E = 9525;
const shape = (over: Partial<DiagramDrawingShape>): DiagramDrawingShape => ({
	modelId: 'm1',
	frame: { x: 0, y: 0, width: 100 * E, height: 50 * E },
	geometry: 'rect',
	has3d: false,
	...over,
});
type SmartArt = HTMLElement & {
	drawing: DiagramDrawing | undefined;
	report?: SmartArtRenderReport;
};
const mount = (): SmartArt => {
	document.body.innerHTML = '<office-ui-smartart label="Process"></office-ui-smartart>';
	return document.body.firstElementChild as SmartArt;
};

describe('office-ui-smartart', () => {
	it('uses the native Excel font reference instead of black shell text', async () => {
		const workbook = await loadXlsx(
			new Uint8Array(
				readFileSync(
					new NodeURL('../../../core/xlsx/__fixtures__/excel-smartart.xlsx', import.meta.url),
				),
			),
		);
		const obj = workbook.sheets[0]!.drawings.find((item) => item.kind === 'smartArt');
		if (!obj || obj.kind !== 'smartArt') throw new Error('Missing native SmartArt fixture');
		const node = document.createElement('div');
		document.body.append(node);
		paintSmartArt(createTestContext(workbook), node, obj, workbook.theme);
		const art = node.querySelector('office-ui-smartart')!;
		const texts = Array.from(art.shadowRoot!.querySelectorAll('text'));
		const native = JSON.parse(
			readFileSync(new NodeURL('./excel-appearance.json', import.meta.url), 'utf8'),
		) as {
			diagrams: {
				nodes: { text: string; fontColor: string; fontName: string; fontSizePt: number }[];
			}[];
		};
		const nodes = native.diagrams[0]!.nodes;
		expect(texts.map((text) => text.textContent)).toEqual(nodes.map((node) => node.text));
		for (const [index, text] of texts.entries()) {
			const expected = nodes[index]!;
			expect(text.getAttribute('fill')).toBe(expected.fontColor);
			expect(Number(text.getAttribute('font-size'))).toBeCloseTo((expected.fontSizePt * 96) / 72);
			expect(text.getAttribute('font-family')).toBe(expected.fontName);
		}
	});

	it('keeps explicit run colors and fonts above theme references and refreshes font changes', () => {
		const art = mount() as SmartArt & {
			schemeColors: Record<string, string>;
			schemeFonts: { major?: string; minor?: string };
		};
		art.schemeColors = { lt1: '#FFFFFF' };
		art.schemeFonts = { minor: 'Aptos Narrow', major: 'Aptos Display' };
		const base = shape({
			style: {
				font: { fontIndex: 'minor', color: { kind: 'scheme', value: 'lt1', transforms: [] } },
			},
			text: {
				text: 'Run',
				paragraphs: [
					{
						runs: [
							{
								text: 'Run',
								typeface: 'Courier New',
								color: { kind: 'srgb', value: 'FF0000', transforms: [] },
							},
						],
					},
				],
			},
		});
		art.drawing = { issues: [], shapes: [base] };
		let text = art.shadowRoot!.querySelector('text')!;
		expect(text.getAttribute('fill')).toBe('#FF0000');
		expect(text.getAttribute('font-family')).toBe('Courier New');
		base.text!.paragraphs[0]!.runs[0]!.typeface = '+mj-lt';
		art.drawing = { issues: [], shapes: [base] };
		expect(art.shadowRoot!.querySelector('text')!.getAttribute('font-family')).toBe(
			'Aptos Display',
		);
		art.schemeFonts = { major: 'Cambria' };
		text = art.shadowRoot!.querySelector('text')!;
		expect(text.getAttribute('font-family')).toBe('Cambria');
	});

	it('reuses DrawingML color transforms for fills and reports unsupported transforms', () => {
		expect(
			colorToCss(
				{ kind: 'scheme', value: 'accent1', transforms: [{ name: 'lumMod', value: '50000' }] },
				{ accent1: '#FF0000' },
			),
		).toBe('#800000');
		expect(
			colorToCss(
				{ kind: 'srgb', value: 'FF0000', transforms: [{ name: 'alpha', value: '50000' }] },
				{},
			),
		).toBe('rgba(255, 0, 0, 0.5)');
		const art = mount();
		art.drawing = {
			issues: [],
			shapes: [
				shape({
					fill: {
						kind: 'solid',
						color: {
							kind: 'srgb',
							value: 'FF0000',
							transforms: [{ name: 'redMod', value: '50000' }],
						},
					},
				}),
			],
		};
		expect(art.report?.unappliedColorTransforms).toEqual(['redMod']);
	});
	it('draws a core DiagramDrawing as an SVG image with an accessible name', () => {
		const el = mount();
		el.drawing = {
			issues: [],
			shapes: [
				shape({
					fill: { kind: 'solid', color: { kind: 'srgb', value: '4472C4', transforms: [] } },
					geometry: 'roundRect',
					text: {
						text: 'Plan',
						paragraphs: [{ runs: [{ text: 'Plan', sizePt: 18, bold: true }] }],
					},
				}),
				shape({
					modelId: 'm2',
					frame: { x: 120 * E, y: 0, width: 50 * E, height: 50 * E },
					geometry: 'ellipse',
				}),
			],
		};
		const svg = el.shadowRoot!.querySelector('svg')!;
		expect(svg.getAttribute('role')).toBe('img');
		expect(svg.getAttribute('aria-label')).toBe('Process');
		expect(svg.getAttribute('viewBox')).toBe('0 0 170 50');
		expect(svg.querySelectorAll('g[data-model-id]')).toHaveLength(2);
		expect(svg.querySelector('rect')!.getAttribute('fill')).toBe('#4472C4');
		expect(svg.querySelector('ellipse')).not.toBeNull();
		expect(svg.querySelector('text')!.textContent).toBe('Plan');
		expect(el.report?.shapeCount).toBe(2);
		expect(el.report?.approximatedGeometries).toEqual([]);
	});

	it('reports approximations and flattened 3D instead of hiding them', () => {
		const el = mount();
		const events: SmartArtRenderReport[] = [];
		el.addEventListener('office-smartart-render', (e) => events.push((e as CustomEvent).detail));
		el.drawing = {
			issues: [],
			shapes: [
				shape({ geometry: 'chevron', has3d: true }),
				shape({
					modelId: 'g',
					fill: {
						kind: 'gradient',
						stops: [{ position: 0, color: { kind: 'srgb', value: 'FF0000', transforms: [] } }],
					},
				}),
			],
		};
		expect(events).toHaveLength(1);
		expect(events[0]).toMatchObject({
			approximatedGeometries: ['chevron'],
			flattened3d: 1,
			approximatedFills: ['gradient'],
		});
		expect(el.shadowRoot!.querySelector('[data-approximated="chevron"]')).not.toBeNull();
		expect(el.shadowRoot!.querySelector('[data-flattened-3d]')).not.toBeNull();
	});

	it('scales custom geometry paths and resolves scheme colours from the host', () => {
		const el = mount() as SmartArt & { schemeColors: Record<string, string> };
		el.schemeColors = { accent1: '#112233' };
		el.drawing = {
			issues: [],
			shapes: [
				shape({
					geometry: 'custom',
					fill: { kind: 'solid', color: { kind: 'scheme', value: 'accent1', transforms: [] } },
					paths: [
						{
							width: 10,
							height: 10,
							commands: [{ op: 'M', x: 0, y: 0 }, { op: 'L', x: 10, y: 10 }, { op: 'Z' }],
						},
					],
				}),
			],
		};
		expect(el.shadowRoot!.querySelector('path')!.getAttribute('d')).toBe('M0 0L100 50Z');
		expect(el.shadowRoot!.querySelector('g > g')!.getAttribute('fill')).toBe('#112233');
	});

	it('says so when there is no cached drawing and renders nothing for undefined', () => {
		const el = mount();
		expect(el.shadowRoot!.querySelector('svg')).toBeNull();
		el.drawing = { shapes: [], issues: [] };
		expect(el.shadowRoot!.querySelector('.empty')!.textContent).toMatch(/no cached drawing/);
		el.drawing = undefined;
		expect(el.shadowRoot!.querySelector('.empty')!.textContent).toBe('');
	});
});
