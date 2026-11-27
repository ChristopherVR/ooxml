/** Page 1: the logical architecture, one container per tier, readings flowing left to right. */
import type { Drawing } from '../lib/drawing';
import { polyline } from '../lib/geometry';
import type { Node, PageBuilder } from '../lib/page';
import { FOG, MIST, PAPER, RUN, SLATE, STONE } from '../lib/palette';
import { SIDE } from '../lib/shape';
import type { Rich } from '../lib/text';
import { CLIENTS, INVENTORY, INVENTORY_RECORDSET, TIERS } from './inventory';
import { CARD, tier } from './stencil';

export const LOGICAL_PAGE = 'Logical architecture';
export const DEPLOYMENT_PAGE = 'Deployment network';

const COLUMNS = [2.0, 5.2, 8.4, 11.6, 14.8] as const;
const ROWS = [7.5, 5.35, 3.2] as const;
const BOX = { halfWidth: 1.3, bottom: 2.4, top: 9.1 };
/** Above the cards and below the container headings: where control traffic crosses the tiers. */
const CONTROL_Y = 8.35;

/** A card's text: the name, and under it a second line. */
const two = (name: string, second: string, field?: string): Rich => [
	{ text: `${name}\n`, style: RUN.name },
	{ text: second, style: RUN.small, ...(field ? { field } : {}) },
];

async function drawTiers(drawing: Drawing, page: PageBuilder): Promise<string[]> {
	const ids = page.reserve(TIERS.length);
	const pageId = page.pageId;
	for (const [index, entry] of TIERS.entries()) {
		const shapeId = ids[index];
		const x = COLUMNS[index];
		if (!shapeId || x === undefined) continue;
		await drawing.edit(
			[
				{
					type: 'insert-container',
					pageId,
					shapeId,
					memberIds: [],
					style: 'plain',
					heading: entry.name,
					box: {
						x,
						y: (BOX.bottom + BOX.top) / 2,
						width: BOX.halfWidth * 2,
						height: BOX.top - BOX.bottom,
					},
				},
			],
			`container ${entry.name}`,
		);
		await drawing.edit(
			[
				{
					type: 'format-shape',
					pageId,
					shapeId,
					fillColor: FOG,
					lineColor: MIST,
					lineWeight: 0.75,
					rounding: 6,
				},
				{
					type: 'format-text',
					pageId,
					shapeId,
					fontFamily: 'Rubik',
					fontSize: 10,
					fontColor: entry.hue.deep,
				},
			],
			`format ${entry.name}`,
		);
		page.assign([entry.name], [shapeId]);
	}
	return ids;
}

type Nodes = Record<string, Node>;

function drawNodes(page: PageBuilder): Nodes {
	const nodes: Nodes = {};
	for (const [key, master, column, row, name, , address] of INVENTORY) {
		const entry = TIERS[column];
		if (!entry) continue;
		nodes[key] = page.drop(master, {
			x: COLUMNS[column],
			y: ROWS[row],
			text: two(name, address, 'Prop.IP'),
			childCells: tier(entry.hue),
			layers: [entry.name],
		});
	}
	const clients = TIERS[4];
	if (!clients) throw new Error('No client tier.');
	for (const [key, master, row, name, note] of CLIENTS) {
		nodes[key] = page.drop(master, {
			x: COLUMNS[4],
			y: ROWS[row],
			text: two(name, note),
			childCells: tier(clients.hue),
			data: { Role: note, Owner: 'Product', Status: 'Live' },
			layers: [clients.name],
		});
	}
	return nodes;
}

function drawFlows(page: PageBuilder, n: Nodes): void {
	const node = (key: string) => {
		const found = n[key];
		if (!found) throw new Error(`No node ${key}`);
		return found;
	};
	const { top, right, bottom, left } = SIDE;
	const data = (from: string, p: number, to: string, q: number, label?: string) =>
		page.connect('Data flow', node(from).port(p), node(to).port(q), {
			layers: ['Data flow'],
			...(label ? { label } : {}),
		});
	// Between two tiers there is room for one short word, set on the line.
	const between = (from: string, p: number, to: string, text: string) =>
		page.connect('Data flow', node(from).port(p), node(to).port(left), {
			layers: ['Data flow'],
			label: { text, offset: 1.6, onLine: true },
		});
	data('garden', bottom, 'hub', top, '868 MHz');
	data('roof', top, 'hub', bottom, '868 MHz');
	between('hub', right, 'gateway', 'MQTT');
	data('gateway', bottom, 'queue', top);
	data('queue', bottom, 'stream', top);
	data('stream', right, 'store', left);
	between('stream', 6, 'archive', 'raw');
	data('store', right, 'forecast', left);
	between('store', 6, 'alerts', 'live');
	data('forecast', top, 'alerts', bottom, 'outlook');
	data('accounts', right, 'api', left, 'accounts');
	data('alerts', top, 'api', bottom, 'alerts');
	data('api', right, 'web', left);
	// The app's feed leaves the API downwards and crosses to the client tier between the rows.
	const feed = node('api').port(8);
	const app = node('app').port(top);
	const feedY = (ROWS[0] + ROWS[1]) / 2;
	page.connect('Data flow', feed, app, {
		layers: ['Data flow'],
		via: [
			{ x: feed.x, y: feedY },
			{ x: app.x, y: feedY },
		],
	});
	data('forecast', right, 'partner', left, 'export');
	const control = { layers: ['Control'] };
	page.connect('Control flow', node('alerts').port(right), node('app').port(left), {
		...control,
		label: 'push',
	});
	const api = node('api').port(top);
	const gateway = node('gateway').port(top);
	page.connect('Control flow', api, gateway, {
		...control,
		via: [
			{ x: api.x, y: CONTROL_Y },
			{ x: gateway.x, y: CONTROL_Y },
		],
		label: { text: 'device commands and firmware', width: 2.1 },
	});
}

/** Notes pointing at three nodes, placed under the tiers; the library draws them as callouts. */
async function drawCallouts(drawing: Drawing, page: PageBuilder, n: Nodes): Promise<void> {
	const notes: [string, number, string][] = [
		[
			'roof',
			0,
			'Stations report every 60 seconds and buffer a day of readings when the hub is away.',
		],
		['store', 2, 'Raw readings are kept for 13 months, hourly roll-ups for five years.'],
		[
			'forecast',
			3,
			'A nowcast per station every 10 minutes, from its own history and its neighbours.',
		],
	];
	const pageId = page.pageId;
	for (const [key, column, text] of notes) {
		const [shapeId, leaderId] = page.reserve(2);
		const target = n[key];
		const x = COLUMNS[column];
		if (!shapeId || !leaderId || !target || x === undefined) continue;
		await drawing.edit(
			[
				{
					type: 'insert-callout',
					pageId,
					shapeId,
					leaderId,
					targetId: String(target.id),
					style: 'rounded',
					text,
					box: { x, y: 1.45, width: 2.6, height: 0.66 },
				},
			],
			`callout ${key}`,
		);
		await drawing.edit(
			[
				{
					type: 'format-shape',
					pageId,
					shapeId,
					fillColor: PAPER,
					lineColor: STONE,
					lineWeight: 0.75,
				},
				{ type: 'format-text', pageId, shapeId, fontSize: 7.5, fontColor: SLATE },
				{ type: 'format-shape', pageId, shapeId: leaderId, lineColor: STONE, lineWeight: 0.75 },
			],
			`format callout ${key}`,
		);
		page.assign(['Annotations'], [shapeId, leaderId]);
	}
}

function drawLegend(page: PageBuilder): void {
	const notes = { layers: ['Annotations'] };
	const left = 4.0;
	const entries: [string, string][] = [
		['Data flow', 'Readings and results'],
		['Control flow', 'Commands and notifications'],
	];
	entries.forEach(([style, caption], index) => {
		const y = 1.62 - index * 0.34;
		page.draw(
			{
				style,
				paths: [
					polyline([
						[0, 0.5],
						[1, 0.5],
					]),
				],
			},
			{ x: left + 0.3, y, w: 0.6, h: 0.1, ...notes },
		);
		page.note('Caption', caption, { left: left + 0.75, y, w: 1.9 });
	});
	page.paragraph(
		'Each tier and each kind of flow is on a layer of its own, so a tier can be hidden or printed alone.',
		{ left: 13.5, top: 1.78, w: 2.6, h: 0.66 },
	);
}

/**
 * Draws the page and links its nodes to the inventory. The address fields are still held back;
 * the caller applies them once no page needs the library's shape data edits any more.
 */
export async function drawLogicalPage(drawing: Drawing): Promise<PageBuilder> {
	const page = drawing.page(LOGICAL_PAGE);
	page.deferFields = true;
	const tiers = await drawTiers(drawing, page);
	const nodes = drawNodes(page);
	drawFlows(page, nodes);
	drawLegend(page);
	await drawing.place(page);
	const members = (column: number) =>
		Object.values(nodes)
			.filter((node) => Math.abs(node.x - (COLUMNS[column] ?? 0)) < CARD.w / 2)
			.map((node) => String(node.id));
	await drawing.edit(
		tiers.map((shapeId, index) => ({
			type: 'set-container-members' as const,
			pageId: page.pageId,
			shapeId,
			memberIds: members(index),
		})),
		'tier membership',
	);
	await drawing.linkData(
		page,
		INVENTORY_RECORDSET,
		INVENTORY.map((row, index) => [nodes[row[0]]?.id ?? 0, index + 1] as const),
	);
	await drawCallouts(drawing, page, nodes);
	// Nothing new to write: this puts the callouts on their layer.
	await drawing.place(page);
	return page;
}
