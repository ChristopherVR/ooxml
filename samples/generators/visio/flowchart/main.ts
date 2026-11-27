/** Page 1: the cross-functional flowchart, from the order to the invoice. */
import type { Drawing } from '../lib/drawing';
import { PAGE } from '../lib/furniture';
import type { Node, PageBuilder } from '../lib/page';
import { HUES, MIST, PAPER, SLATE } from '../lib/palette';
import { SIDE } from '../lib/shape';
import { drawLegend } from './legend';
import { REGISTER, REGISTER_RECORDSET, registered } from './register';

export const MAIN_PAGE = 'Order to delivery';
export const RETURNS_PAGE = 'Returns';

const LANES = ['Customer', 'Sales', 'Warehouse', 'Carrier', 'Finance'] as const;
type Lane = (typeof LANES)[number];

/** The swimlane frame: title bar on top, a heading band on the left of each lane. */
const FRAME = { left: PAGE.left, right: 14.0, bottom: 0.9, top: 9.5, title: 0.35, band: 0.5 };
const LANE_HEIGHT = (FRAME.top - FRAME.title - FRAME.bottom) / LANES.length;
const laneY = (lane: Lane) => FRAME.top - FRAME.title - (LANES.indexOf(lane) + 0.5) * LANE_HEIGHT;
/** Seven columns across the lanes. */
const columnX = (column: number) => FRAME.left + FRAME.band + 0.95 + (column - 1) * 1.8;
/** Phase boundaries sit halfway between two columns. */
const PHASES = [
	{ name: 'Order', from: FRAME.left + FRAME.band },
	{ name: 'Fulfil', from: (columnX(2) + columnX(3)) / 2 },
	{ name: 'Deliver and settle', from: (columnX(5) + columnX(6)) / 2 },
];

async function drawLanes(drawing: Drawing, page: PageBuilder): Promise<string[]> {
	const [frame, list, phases, ...rest] = page.reserve(3 + LANES.length + PHASES.length - 1);
	const lanes = rest.slice(0, LANES.length);
	const separators = rest.slice(LANES.length);
	if (!frame || !list || !phases) throw new Error('No shape IDs for the swimlanes.');
	const pageId = page.pageId;
	await drawing.edit(
		[
			{
				type: 'create-swimlanes',
				pageId,
				frameId: frame,
				laneIds: lanes,
				headings: [...LANES],
				box: {
					x: (FRAME.left + FRAME.right) / 2,
					y: (FRAME.bottom + FRAME.top) / 2,
					width: FRAME.right - FRAME.left,
					height: FRAME.top - FRAME.bottom,
				},
				direction: 'horizontal',
				title: 'Kestrelline Outfitters: from order to delivery',
				style: 3,
				listId: list,
				phaseListId: phases,
			},
		],
		'swimlanes',
	);
	for (const [index, shapeId] of separators.entries()) {
		const phase = PHASES[index + 1];
		if (!phase) continue;
		await drawing.edit(
			[
				{
					type: 'insert-separator',
					pageId,
					frameId: frame,
					shapeId,
					at: phase.from,
					label: phase.name,
				},
			],
			`phase ${phase.name}`,
		);
	}
	// The library draws the lanes in its own colours; its formatting commands restyle them.
	const format = { pageId, lineColor: MIST, lineWeight: 0.75 } as const;
	const text = { pageId, fontFamily: 'Rubik' } as const;
	await drawing.edit(
		[
			{ type: 'format-shape', ...format, shapeId: frame, fillColor: HUES.sea.deep },
			{ type: 'format-text', ...text, shapeId: frame, fontColor: PAPER, fontSize: 10 },
			...lanes.flatMap((shapeId) => [
				{ type: 'format-shape' as const, ...format, shapeId, fillColor: PAPER },
				{ type: 'format-text' as const, ...text, shapeId, fontColor: HUES.sea.deep, fontSize: 9 },
			]),
			...separators.flatMap((shapeId) => [
				{ type: 'format-shape' as const, pageId, shapeId, fillColor: MIST, lineColor: MIST },
				{ type: 'format-text' as const, ...text, shapeId, fontColor: SLATE, fontSize: 8 },
			]),
		],
		'swimlane formatting',
	);
	page.assign(['Swimlanes'], [frame, ...lanes, ...separators]);
	return lanes;
}

interface Steps {
	[key: string]: Node;
}

function drawSteps(page: PageBuilder): Steps {
	const at = (column: number, lane: Lane) => ({
		x: columnX(column),
		y: laneY(lane),
		layers: ['Steps'],
	});
	const role = (owner: string, system: string) => ({ Owner: owner, System: system });
	return {
		place: page.drop('Start or end', {
			...at(1, 'Customer'),
			text: 'Order placed',
			data: role('Customer', 'Storefront'),
		}),
		validate: page.drop('Step', { ...at(1, 'Sales'), text: registered('S-01') }),
		authorise: page.drop('Step', { ...at(1, 'Finance'), text: registered('S-02') }),
		paid: page.drop('Decision', { ...at(2, 'Sales'), text: registered('S-03', false) }),
		declined: page.drop('Start or end', {
			...at(2, 'Customer'),
			text: 'Order declined',
			data: role('Customer', 'Storefront'),
		}),
		stock: page.drop('Data', { ...at(3, 'Warehouse'), text: registered('S-04', false) }),
		inStock: page.drop('Decision', { ...at(4, 'Warehouse'), text: registered('S-05', false) }),
		backOrder: page.drop('Step', { ...at(4, 'Sales'), text: registered('S-06') }),
		notice: page.drop('Document', {
			...at(4, 'Customer'),
			text: 'Back-order notice',
			data: role('Sales desk', 'Storefront'),
		}),
		pick: page.drop('Step', { ...at(5, 'Warehouse'), text: registered('S-07') }),
		label: page.drop('Document', {
			...at(5, 'Carrier'),
			text: 'Shipping label',
			data: role('Warehouse', 'Dispatch board'),
		}),
		collect: page.drop('Step', { ...at(6, 'Carrier'), text: registered('S-08') }),
		delivered: page.drop('Decision', { ...at(7, 'Carrier'), text: registered('S-09', false) }),
		received: page.drop('Start or end', {
			...at(7, 'Customer'),
			text: 'Parcel received',
			data: role('Customer', 'Parcel tracking'),
		}),
		returns: page.drop('Off-page reference', {
			...at(6, 'Customer'),
			text: 'Returns',
			link: { page: RETURNS_PAGE, description: 'Go to the returns sub-process' },
		}),
		capture: page.drop('Step', { ...at(6, 'Finance'), text: registered('S-10') }),
		invoice: page.drop('Document', {
			...at(7, 'Finance'),
			text: 'Invoice',
			data: role('Finance', 'Payments ledger'),
		}),
	};
}

function drawFlows(page: PageBuilder, s: Steps): void {
	const flows = { layers: ['Flows'] };
	const seq = (from: Node, a: number, to: Node, b: number, label?: string) =>
		page.connect('Sequence flow', from.port(a), to.port(b), {
			...flows,
			...(label ? { label: { text: label, offset: 0.2 } } : {}),
		});
	const msg = (from: Node, a: number, to: Node, b: number) =>
		page.connect('Message flow', from.port(a), to.port(b), flows);
	const need = (name: string) => {
		const node = s[name];
		if (!node) throw new Error(`No step ${name}`);
		return node;
	};
	const [place, validate, authorise, paid, declined, stock, inStock, backOrder, notice] = [
		'place',
		'validate',
		'authorise',
		'paid',
		'declined',
		'stock',
		'inStock',
		'backOrder',
		'notice',
	].map(need) as [Node, Node, Node, Node, Node, Node, Node, Node, Node];
	const [pick, label, collect, delivered, received, returns, capture, invoice] = [
		'pick',
		'label',
		'collect',
		'delivered',
		'received',
		'returns',
		'capture',
		'invoice',
	].map(need) as [Node, Node, Node, Node, Node, Node, Node, Node];
	const { top, right, bottom, left } = SIDE;
	seq(place, bottom, validate, top);
	seq(validate, bottom, authorise, top);
	seq(authorise, right, paid, bottom);
	seq(paid, top, declined, bottom, 'No');
	seq(paid, right, stock, top, 'Yes');
	seq(stock, right, inStock, left);
	seq(inStock, top, backOrder, bottom, 'No');
	msg(backOrder, top, notice, bottom);
	page.connect('Sequence flow', backOrder.port(right), pick.port(top), {
		...flows,
		label: { text: 'Restocked', offset: 0.45 },
	});
	seq(inStock, right, pick, left, 'Yes');
	seq(pick, bottom, label, top);
	seq(label, right, collect, left);
	seq(collect, right, delivered, left);
	seq(delivered, top, received, bottom, 'Yes');
	// A failed delivery loops back under the two shapes and is tried again the next day.
	const loopY = delivered.bottom - 0.26;
	const retry = collect.port(8);
	page.connect('Sequence flow', delivered.port(bottom), retry, {
		...flows,
		via: [
			{ x: delivered.x, y: loopY },
			{ x: retry.x, y: loopY },
		],
		label: {
			text: 'No: retry next day',
			offset: 0.26 + (delivered.x - retry.x) / 2,
			width: 1.25,
			onLine: true,
		},
	});
	seq(collect, 9, capture, 4);
	msg(capture, right, invoice, left);
	msg(received, left, returns, right);
}

/**
 * Draws the page and links its steps to the register. The steps' text fields are still held
 * back: the caller applies them once no page needs the library's shape data edits any more.
 */
export async function drawMainPage(drawing: Drawing): Promise<PageBuilder> {
	const page = drawing.page(MAIN_PAGE);
	page.deferFields = true;
	const lanes = await drawLanes(drawing, page);
	page.note('Phase label', PHASES[0]?.name ?? '', {
		left: FRAME.left + FRAME.band + 0.1,
		y: FRAME.top - FRAME.title - 0.15,
		w: 2,
		h: 0.2,
	});
	const steps = drawSteps(page);
	drawFlows(page, steps);
	drawLegend(page);
	await drawing.place(page);
	// Each step joins the lane it sits in, so moving or resizing a lane carries its steps.
	await drawing.edit(
		lanes.map((shapeId, index) => ({
			type: 'set-container-members' as const,
			pageId: page.pageId,
			shapeId,
			memberIds: Object.values(steps)
				.filter((node) => Math.floor((FRAME.top - FRAME.title - node.y) / LANE_HEIGHT) === index)
				.map((node) => String(node.id)),
		})),
		'lane membership',
	);
	const row = (step: string) => REGISTER.findIndex((entry) => entry[0] === step) + 1;
	const linked: [string, string][] = [
		['validate', 'S-01'],
		['authorise', 'S-02'],
		['paid', 'S-03'],
		['stock', 'S-04'],
		['inStock', 'S-05'],
		['backOrder', 'S-06'],
		['pick', 'S-07'],
		['collect', 'S-08'],
		['delivered', 'S-09'],
		['capture', 'S-10'],
	];
	await drawing.linkData(
		page,
		REGISTER_RECORDSET,
		linked.map(([name, step]) => [steps[name]?.id ?? 0, row(step)] as const),
	);
	return page;
}
