/** Page 2: the returns sub-process, drawn as three containers that are its stages. */
import type { Drawing } from '../lib/drawing';
import type { Node, PageBuilder } from '../lib/page';
import { FOG, HUES, MIST } from '../lib/palette';
import { SIDE } from '../lib/shape';
import { MAIN_PAGE, RETURNS_PAGE } from './main';

const ROWS = [7.95, 6.25, 4.55, 2.85] as const;
const STAGES = [
	{ heading: '1  Request', left: 1.0, right: 5.8 },
	{ heading: '2  Inspect', left: 6.1, right: 10.9 },
	{ heading: '3  Settle', left: 11.2, right: 16.0 },
] as const;
const BOX = { bottom: 2.05, top: 9.0 };

/** The library's containers, restyled with its own formatting commands. */
async function drawStages(drawing: Drawing, page: PageBuilder): Promise<string[]> {
	const ids = page.reserve(STAGES.length);
	const pageId = page.pageId;
	for (const [index, stage] of STAGES.entries()) {
		const shapeId = ids[index];
		if (!shapeId) continue;
		await drawing.edit(
			[
				{
					type: 'insert-container',
					pageId,
					shapeId,
					memberIds: [],
					style: 'plain',
					heading: stage.heading,
					box: {
						x: (stage.left + stage.right) / 2,
						y: (BOX.bottom + BOX.top) / 2,
						width: stage.right - stage.left,
						height: BOX.top - BOX.bottom,
					},
				},
			],
			`container ${stage.heading}`,
		);
	}
	await drawing.edit(
		ids.flatMap((shapeId) => [
			{
				type: 'format-shape' as const,
				pageId,
				shapeId,
				fillColor: FOG,
				lineColor: MIST,
				lineWeight: 0.75,
				rounding: 6,
			},
			{
				type: 'format-text' as const,
				pageId,
				shapeId,
				fontFamily: 'Rubik',
				fontSize: 10,
				fontColor: HUES.sea.deep,
			},
		]),
		'container formatting',
	);
	page.assign(['Stages'], ids);
	return ids;
}

export async function drawReturnsPage(drawing: Drawing): Promise<void> {
	const page = drawing.page(RETURNS_PAGE);
	const stages = await drawStages(drawing, page);
	const steps = { layers: ['Steps'] };
	const role = (owner: string, system: string, status = 'In place') => ({
		Owner: owner,
		System: system,
		Status: status,
	});
	const [a, b, c] = [2.2, 7.3, 12.4];
	const second = 2.4;
	const members: Node[][] = [[], [], []];
	const put = (stage: number, node: Node) => {
		members[stage]?.push(node);
		return node;
	};

	const from = put(
		0,
		page.drop('Off-page reference', {
			x: a,
			y: ROWS[0],
			text: 'From delivery',
			link: { page: MAIN_PAGE, description: 'Back to the order-to-delivery chart' },
			...steps,
		}),
	);
	const log = put(
		0,
		page.drop('Step', {
			x: a,
			y: ROWS[1],
			text: 'Log return request',
			data: { ...role('Customer care', 'Service desk'), Step: 'R-01', Duration: 0.2, Cost: 0.6 },
			...steps,
		}),
	);
	const inTime = put(
		0,
		page.drop('Decision', {
			x: a,
			y: ROWS[2],
			text: 'Within 30 days?',
			data: { ...role('Customer care', 'Service desk'), Step: 'R-02' },
			...steps,
		}),
	);
	const declined = put(
		0,
		page.drop('Start or end', {
			x: a + second,
			y: ROWS[2],
			text: 'Return declined',
			data: role('Customer care', 'Service desk'),
			...steps,
		}),
	);
	const label = put(
		0,
		page.drop('Document', {
			x: a,
			y: ROWS[3],
			text: 'Prepaid return label',
			data: role('Customer care', 'Dispatch board'),
			...steps,
		}),
	);
	const receive = put(
		1,
		page.drop('Step', {
			x: b,
			y: ROWS[3],
			text: 'Receive parcel',
			data: { ...role('Warehouse', 'Returns bench'), Step: 'R-03', Duration: 0.1, Cost: 0.4 },
			...steps,
		}),
	);
	const inspect = put(
		1,
		page.drop('Step', {
			x: b,
			y: ROWS[2],
			text: 'Inspect item',
			data: { ...role('Warehouse', 'Returns bench'), Step: 'R-04', Duration: 0.3, Cost: 1.9 },
			...steps,
		}),
	);
	const resale = put(
		1,
		page.drop('Decision', {
			x: b,
			y: ROWS[1],
			text: 'Fit for resale?',
			data: { ...role('Warehouse', 'Returns bench'), Step: 'R-05' },
			...steps,
		}),
	);
	const restock = put(
		1,
		page.drop('Step', {
			x: b,
			y: ROWS[0],
			text: 'Return to stock',
			data: { ...role('Warehouse', 'Stock ledger'), Step: 'R-06', Duration: 0.2, Cost: 0.7 },
			...steps,
		}),
	);
	const writeOff = put(
		1,
		page.drop('Step', {
			x: b + second,
			y: ROWS[1],
			text: 'Write off and recycle',
			data: {
				...role('Warehouse', 'Stock ledger', 'Under review'),
				Step: 'R-07',
				Duration: 0.2,
				Cost: 3.2,
			},
			...steps,
		}),
	);
	const refund = put(
		2,
		page.drop('Step', {
			x: c,
			y: ROWS[1],
			text: 'Refund customer',
			data: { ...role('Finance', 'Payments ledger'), Step: 'R-08', Duration: 0.1, Cost: 0.3 },
			...steps,
		}),
	);
	const ledger = put(
		2,
		page.drop('Data', {
			x: c,
			y: ROWS[2],
			text: 'Payments ledger',
			data: role('Finance', 'Payments ledger'),
			...steps,
		}),
	);
	const credit = put(
		2,
		page.drop('Document', {
			x: c + second,
			y: ROWS[2],
			text: 'Credit note',
			data: role('Finance', 'Payments ledger'),
			...steps,
		}),
	);
	const closed = put(
		2,
		page.drop('Start or end', {
			x: c,
			y: ROWS[3],
			text: 'Return closed',
			data: role('Customer care', 'Service desk'),
			...steps,
		}),
	);

	const flows = { layers: ['Flows'] };
	const { top, right, bottom, left } = SIDE;
	const seq = (source: Node, p: number, target: Node, q: number, text?: string) =>
		page.connect('Sequence flow', source.port(p), target.port(q), {
			...flows,
			...(text ? { label: { text, offset: 0.2 } } : {}),
		});
	seq(from, bottom, log, top);
	seq(log, bottom, inTime, top);
	seq(inTime, right, declined, left, 'No');
	seq(inTime, bottom, label, top, 'Yes');
	page.connect('Message flow', label.port(right), receive.port(left), flows);
	seq(receive, top, inspect, bottom);
	seq(inspect, top, resale, bottom);
	seq(resale, top, restock, bottom, 'Yes');
	seq(resale, right, writeOff, left, 'No');
	seq(restock, right, refund, top);
	seq(writeOff, right, refund, left);
	seq(refund, bottom, ledger, top);
	page.connect('Message flow', ledger.port(right), credit.port(left), flows);
	seq(ledger, bottom, closed, top);

	page.note('Heading', 'Returns sub-process', { left: 1.0, y: 9.33, w: 6 });
	page.paragraph(
		'Starts when a customer asks to send a parcel back. Each stage is a container: move it and its steps move with it.',
		{ left: 1.0, top: 1.8, w: 7.5, h: 0.4 },
	);
	page.paragraph(
		'Refunds reach the original payment method within five working days of inspection.',
		{ left: 13.2, top: 7.9, w: 2.6, h: 0.7 },
	);
	await drawing.place(page);
	await drawing.edit(
		stages.map((shapeId, index) => ({
			type: 'set-container-members' as const,
			pageId: page.pageId,
			shapeId,
			memberIds: (members[index] ?? []).map((node) => String(node.id)),
		})),
		'stage membership',
	);
}
