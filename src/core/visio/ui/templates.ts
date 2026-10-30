import { createVsdx, editVsdx, type VisioEdit } from '../index';
import { createOrgChartTemplate } from './org-chart-template';
import { shapeEdits, type SampleBox, type SampleLink, type SamplePage } from './sample-drawing';

/** The templates File > New offers, in Visio's order: Blank Drawing first. */
export const VISIO_TEMPLATES = [
	{ id: 'blank', name: 'Blank Drawing', description: 'Start an empty drawing.' },
	{
		id: 'basic-flowchart',
		name: 'Basic Flowchart',
		description: 'A short process with a decision, using Basic Flowchart Shapes.',
	},
	{
		id: 'cross-functional-flowchart',
		name: 'Cross-Functional Flowchart',
		description: 'A process across three horizontal swimlanes that stack and resize together.',
	},
	{
		id: 'organization-chart',
		name: 'Basic Organization Chart',
		description: 'A small reporting hierarchy joined by connectors.',
	},
] as const;
export type VisioTemplateId = (typeof VISIO_TEMPLATES)[number]['id'];

export const isVisioTemplateId = (value: unknown): value is VisioTemplateId =>
	VISIO_TEMPLATES.some((template) => template.id === value);

const step = (
	id: string,
	text: string,
	x: number,
	y: number,
	fill = '#D8EFEB',
	master = 'flowchart-process',
	width = 1.7,
	height = 0.7,
): SampleBox => ({ id, text, x, y, width, height, fill, master });
const links = (...pairs: [string, string][]): SampleLink[] =>
	pairs.map(([from, to], index) => ({ id: String(100 + index), from, to }));

const FLOW: SamplePage = {
	id: '0',
	name: 'Basic Flowchart',
	width: 8.5,
	height: 7,
	boxes: [
		step('1', 'Start', 4.25, 6.1, '#F3EEE1', 'flowchart-terminator', 2, 0.6),
		step('2', 'Process step', 4.25, 4.9, '#D8EFEB', 'flowchart-process', 2.4),
		step('3', 'Decision?', 4.25, 3.5, '#DFEBF7', 'flowchart-decision', 1.9, 1.1),
		step('4', 'Next step', 4.25, 2.1, '#D8EFEB', 'flowchart-process', 2.4),
		step('5', 'End', 4.25, 0.9, '#F3EEE1', 'flowchart-terminator', 2, 0.6),
	],
	links: links(['1', '2'], ['2', '3'], ['3', '4'], ['4', '5']),
};

const LANES = ['Customer', 'Sales', 'Fulfilment'];
const SWIM: SamplePage = {
	id: '0',
	name: 'Cross-Functional Flowchart',
	width: 11,
	height: 8.5,
	boxes: [
		step('1', 'Place order', 2.6, 5.6, '#F3EEE1', 'flowchart-terminator', 1.6, 0.6),
		step('2', 'Review order', 4.8, 3.8, '#D8EFEB'),
		step('3', 'In stock?', 7.0, 3.8, '#DFEBF7', 'flowchart-decision', 1.5, 1),
		step('4', 'Ship order', 7.0, 2.0, '#D8EFEB'),
	],
	links: links(['1', '2'], ['2', '3'], ['3', '4']),
};

const SWIM_STEPS = ['1', '2', '3', '4'];
/** Three 1.8 inch lanes under a title bar, 8.4 inches wide, from the lower-left corner (0.6, 1.1). */
const swimlaneFrame = (pageId: string): VisioEdit => ({
	type: 'create-swimlanes',
	pageId,
	frameId: '50',
	laneIds: ['10', '11', '12'],
	headings: LANES,
	box: { x: 4.8, y: 1.1 + (5.4 + 0.35) / 2, width: 8.4, height: 5.4 + 0.35 },
	direction: 'horizontal',
	title: 'Cross-Functional Flowchart',
});

/**
 * A starter drawing as a real VSDX package, built by the editor's own edit API from built-in
 * stencil masters, so every shape can be selected and edited and every connector is glued.
 */
export async function createTemplateVsdx(id: VisioTemplateId): Promise<Uint8Array> {
	if (!isVisioTemplateId(id)) throw new Error('Unknown drawing template.');
	if (id === 'blank') return createVsdx();
	if (id === 'organization-chart') return createOrgChartTemplate();
	const page = id === 'basic-flowchart' ? FLOW : SWIM;
	const blank = await createVsdx({
		width: page.width,
		height: page.height,
		stencils: ['basic-flowchart'],
	});
	const named = await editVsdx(blank, [{ type: 'rename-page', pageId: page.id, name: page.name }]);
	if (page === SWIM) {
		// The frame and its lanes go in first so they sit behind the process shapes, which then
		// join the lane they were drawn in.
		const lanes = await editVsdx(named.bytes, [swimlaneFrame(page.id)]);
		const drawn = await editVsdx(lanes.bytes, shapeEdits(page));
		const joined = [
			{ type: 'reconcile-containers' as const, pageId: page.id, shapeIds: SWIM_STEPS },
		];
		return (await editVsdx(drawn.bytes, joined)).bytes;
	}
	return (await editVsdx(named.bytes, shapeEdits(page))).bytes;
}
