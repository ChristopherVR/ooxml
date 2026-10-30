import { createVsdx, editVsdx, type VisioEdit } from '../index';
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
		description: 'A process across three horizontal swimlanes inside a container.',
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

const LANE_LEFT = 0.6;
const LANE_WIDTH = 8.4;
const LANE_HEIGHT = 1.8;
const LANES = ['Customer', 'Sales', 'Fulfilment'];
/** Lane `index` (0 at the top) as a header and a body, ordinary shapes the container frames. */
function lane(index: number): SampleBox[] {
	const y = 5.6 - index * LANE_HEIGHT;
	const header = 0.5;
	return [
		{
			id: String(10 + index * 2),
			text: LANES[index]!,
			x: LANE_LEFT + header / 2,
			y,
			width: header,
			height: LANE_HEIGHT,
			fill: '#DFEBF7',
		},
		{
			id: String(11 + index * 2),
			text: '',
			x: LANE_LEFT + header + (LANE_WIDTH - header) / 2,
			y,
			width: LANE_WIDTH - header,
			height: LANE_HEIGHT,
			fill: '#FFFFFF',
		},
	];
}
const SWIM: SamplePage = {
	id: '0',
	name: 'Cross-Functional Flowchart',
	width: 11,
	height: 8.5,
	boxes: [
		...LANES.flatMap((_, index) => lane(index)),
		step('1', 'Place order', 2.6, 5.6, '#F3EEE1', 'flowchart-terminator', 1.6, 0.6),
		step('2', 'Review order', 4.8, 3.8, '#D8EFEB'),
		step('3', 'In stock?', 7.0, 3.8, '#DFEBF7', 'flowchart-decision', 1.5, 1),
		step('4', 'Ship order', 7.0, 2.0, '#D8EFEB'),
	],
	links: links(['1', '2'], ['2', '3'], ['3', '4']),
};
const SWIM_MEMBERS = ['10', '11', '12', '13', '14', '15', '1', '2', '3', '4'];

const ORG: SamplePage = {
	id: '0',
	name: 'Organization Chart',
	width: 11,
	height: 8.5,
	boxes: [
		step('1', 'Chief Executive\nName', 5.5, 7, '#DFEBF7', 'rectangle', 2.2, 0.8),
		step('2', 'Manager\nName', 2.6, 5.2, '#D8EFEB', 'rectangle', 2.2, 0.8),
		step('3', 'Manager\nName', 5.5, 5.2, '#D8EFEB', 'rectangle', 2.2, 0.8),
		step('4', 'Manager\nName', 8.4, 5.2, '#D8EFEB', 'rectangle', 2.2, 0.8),
		step('5', 'Position\nName', 1.6, 3.4, '#F3EEE1', 'rectangle', 1.8, 0.8),
		step('6', 'Position\nName', 3.7, 3.4, '#F3EEE1', 'rectangle', 1.8, 0.8),
	],
	links: links(['1', '2'], ['1', '3'], ['1', '4'], ['2', '5'], ['2', '6']),
};

/**
 * A starter drawing as a real VSDX package, built by the editor's own edit API from built-in
 * stencil masters, so every shape can be selected and edited and every connector is glued.
 */
export async function createTemplateVsdx(id: VisioTemplateId): Promise<Uint8Array> {
	if (!isVisioTemplateId(id)) throw new Error('Unknown drawing template.');
	if (id === 'blank') return createVsdx();
	const page = id === 'basic-flowchart' ? FLOW : id === 'organization-chart' ? ORG : SWIM;
	const blank = await createVsdx({
		width: page.width,
		height: page.height,
		stencils: id === 'organization-chart' ? ['basic'] : ['basic-flowchart'],
	});
	const named = await editVsdx(blank, [{ type: 'rename-page', pageId: page.id, name: page.name }]);
	const built = await editVsdx(named.bytes, shapeEdits(page));
	if (page !== SWIM) return built.bytes;
	// Swimlanes: a container frames the lanes and the process shapes (lanes stay ordinary shapes).
	const frame: VisioEdit = {
		type: 'insert-container',
		pageId: page.id,
		shapeId: '50',
		memberIds: SWIM_MEMBERS,
		style: 'plain',
		heading: 'Cross-Functional Flowchart',
	};
	return (await editVsdx(built.bytes, [frame])).bytes;
}
