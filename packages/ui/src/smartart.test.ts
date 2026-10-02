import type { DiagramDrawing, DiagramDrawingShape } from 'ooxml-core/diagram';
import { registerOfficeUi, type SmartArtRenderReport } from './index.js';

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
