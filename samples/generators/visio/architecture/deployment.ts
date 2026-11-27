/** Page 2: the deployment network, one container per network, links between hosts. */
import type { Drawing } from '../lib/drawing';
import { polyline } from '../lib/geometry';
import type { Node, PageBuilder } from '../lib/page';
import { type Hue, FOG, HUES, MIST, RUN } from '../lib/palette';
import { SIDE } from '../lib/shape';
import { DEPLOYMENT_PAGE, LOGICAL_PAGE } from './logical';
import { tier } from './stencil';

const COLUMNS = [2.6, 5.4, 8.2, 11.0] as const;
/** Card centres of the four bands, top to bottom. */
const BANDS = [8.05, 6.0, 3.95, 1.9] as const;
const BAND = { above: 1.0, below: 0.62, side: 1.35 };

interface Network {
	name: string;
	range: string;
	hue: Hue;
	/** First and last column the container spans. */
	span: readonly [number, number];
}
const NETWORKS: readonly Network[] = [
	{ name: 'Customer home', range: '192.168.1.0/24', hue: HUES.teal, span: [0, 2] },
	{ name: 'Edge subnet', range: '10.20.1.0/24', hue: HUES.sea, span: [0, 3] },
	{ name: 'Application subnet', range: '10.20.2.0/24', hue: HUES.amber, span: [0, 3] },
	{ name: 'Data subnet', range: '10.20.3.0/24', hue: HUES.plum, span: [0, 3] },
];

/** [key, master, band, column, name, address, role, monthly cost] */
const HOSTS = [
	['station', 'Weather station', 0, 0, 'Aerie station', '192.168.1.21', 'Outdoor sensor array', 0],
	['hub', 'Home hub', 0, 1, 'Aerie hub', '192.168.1.20', 'Relay for one home', 0],
	['router', 'Gateway', 0, 2, 'Home router', '192.168.1.1', 'Customer equipment', 0],
	['gw2', 'Gateway', 1, 0, 'gw-02', '10.20.1.12', 'MQTT ingress', 310],
	['gw1', 'Gateway', 1, 1, 'gw-01', '10.20.1.11', 'MQTT ingress', 310],
	['balancer', 'Load balancer', 1, 2, 'lb-01', '10.20.1.5', 'TLS termination', 140],
	['firewall', 'Firewall', 1, 3, 'fw-01', '10.20.1.1', 'Perimeter firewall', 220],
	['broker', 'Queue', 2, 0, 'broker-01', '10.20.2.21', 'Telemetry queue', 420],
	['stream', 'Server', 2, 1, 'stream-01', '10.20.2.31', 'Stream processor', 560],
	['api1', 'Server', 2, 2, 'api-01', '10.20.2.11', 'Public API', 190],
	['api2', 'Server', 2, 3, 'api-02', '10.20.2.12', 'Public API', 190],
	['archive', 'Archive', 3, 0, 'archive', '10.20.3.40', 'Object archive', 85],
	['primary', 'Database', 3, 1, 'tsdb-01', '10.20.3.11', 'Telemetry store, primary', 940],
	['replica', 'Database', 3, 2, 'tsdb-02', '10.20.3.12', 'Telemetry store, replica', 610],
	['accounts', 'Database', 3, 3, 'accounts-db', '10.20.3.13', 'Account database', 190],
] as const;

async function drawNetworks(drawing: Drawing, page: PageBuilder): Promise<string[]> {
	const ids = page.reserve(NETWORKS.length);
	const pageId = page.pageId;
	for (const [index, network] of NETWORKS.entries()) {
		const shapeId = ids[index];
		const y = BANDS[index];
		const from = COLUMNS[network.span[0]];
		const to = COLUMNS[network.span[1]];
		if (!shapeId || y === undefined || from === undefined || to === undefined) continue;
		const left = from - BAND.side;
		const right = to + BAND.side;
		await drawing.edit(
			[
				{
					type: 'insert-container',
					pageId,
					shapeId,
					memberIds: [],
					style: 'plain',
					heading: network.name,
					box: {
						x: (left + right) / 2,
						y: y + (BAND.above - BAND.below) / 2,
						width: right - left,
						height: BAND.above + BAND.below,
					},
				},
			],
			`container ${network.name}`,
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
					fontSize: 9.5,
					fontColor: network.hue.deep,
				},
			],
			`format ${network.name}`,
		);
		page.assign([network.name], [shapeId]);
		// The address range sits at the right end of the heading, clear of the links.
		page.note('Fact right', network.range, {
			left: right - 1.15,
			y: y + BAND.above - 0.19,
			w: 1.05,
			h: 0.2,
		});
	}
	return ids;
}

type Hosts = Record<string, Node>;

function drawHosts(page: PageBuilder): Hosts {
	const hosts: Hosts = {};
	for (const [key, master, band, column, name, address, role, cost] of HOSTS) {
		const network = NETWORKS[band];
		if (!network) continue;
		hosts[key] = page.drop(master, {
			x: COLUMNS[column],
			y: BANDS[band],
			text: [
				{ text: `${name}\n`, style: RUN.name },
				{ text: address, field: 'Prop.IP', style: RUN.small },
			],
			childCells: tier(network.hue),
			data: {
				Hostname: band === 0 ? `${key}.home` : `${name}.nimbuswright.example`,
				IP: address,
				Role: role,
				Owner: band === 0 ? 'Customer' : 'Platform',
				Cost: cost,
			},
			layers: [network.name],
		});
	}
	// The internet is no one's network: a card beside the home, above the firewall.
	hosts.internet = page.drop('Network cloud', {
		x: COLUMNS[3],
		y: BANDS[0],
		text: [
			{ text: 'Internet\n', style: RUN.name },
			{ text: 'TLS 1.3, port 8883', style: RUN.small },
		],
		childCells: tier(HUES.coral),
		data: { Role: 'Public network', Owner: 'None' },
		layers: ['Links'],
	});
	return hosts;
}

function drawLinks(page: PageBuilder, h: Hosts): void {
	const host = (key: string) => {
		const found = h[key];
		if (!found) throw new Error(`No host ${key}`);
		return found;
	};
	const { top, right, bottom, left } = SIDE;
	const link = (from: string, p: number, to: string, q: number, label?: string) =>
		page.connect('Network link', host(from).port(p), host(to).port(q), {
			layers: ['Links'],
			...(label ? { label } : {}),
		});
	link('station', right, 'hub', left, '868 MHz');
	link('hub', right, 'router', left, 'Wi-Fi');
	link('router', right, 'internet', left);
	link('internet', bottom, 'firewall', top);
	link('firewall', left, 'balancer', right);
	link('balancer', left, 'gw1', right);
	link('gw1', left, 'gw2', right, 'cluster');
	link('gw2', bottom, 'broker', top);
	link('gw1', 9, 'broker', 5);
	link('broker', right, 'stream', left);
	link('balancer', bottom, 'api1', top);
	link('balancer', 8, 'api2', 4);
	link('stream', bottom, 'primary', top);
	link('stream', 9, 'archive', 5);
	link('api1', bottom, 'replica', top);
	link('api2', bottom, 'accounts', top);
	page.connect('Replication', host('primary').port(right), host('replica').port(left), {
		layers: ['Links'],
		label: 'replica',
	});
}

function drawSidebar(page: PageBuilder): void {
	const left = 13.0;
	const notes = { layers: ['Annotations'] };
	page.note('Heading', 'Address plan', { left, y: 9.33, w: 3.3 });
	NETWORKS.forEach((network, index) => {
		const y = 8.85 - index * 0.42;
		page.note('Fact', network.range, { left, y, w: 1.4 });
		page.note('Caption', network.name, { left: left + 1.45, y, w: 1.9 });
	});
	page.block('Hairline', { left, bottom: 7.2, w: 3.3, h: 0.012 });
	page.note('Heading', 'Links', { left, y: 6.92, w: 3.3 });
	const entries: [string, string][] = [
		['Network link', 'Network connection'],
		['Replication', 'Streaming replication'],
	];
	entries.forEach(([style, caption], index) => {
		const y = 6.48 - index * 0.36;
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
		page.note('Caption', caption, { left: left + 0.75, y, w: 2.5 });
	});
	page.block('Hairline', { left, bottom: 5.6, w: 3.3, h: 0.012 });
	page.note('Heading', 'Notes', { left, y: 5.32, w: 3.3 });
	page.paragraph(
		'Only the edge subnet is reachable from the internet, and only through the firewall. The application and data subnets have no public addresses.',
		{ left, top: 5.1, w: 3.3, h: 0.45 },
	);
	page.paragraph(
		'Every host card carries its address, role, owner and monthly cost as shape data. The second line of a card is a field that shows the address.',
		{ left, top: 4.5, w: 3.3, h: 0.45 },
	);
	page.note('Caption', 'See the logical architecture', {
		left,
		y: 3.82,
		w: 3.3,
		link: { page: LOGICAL_PAGE, description: 'Go to the logical architecture' },
	});
}

/** Draws the page; its address fields are held back like those of the logical page. */
export async function drawDeploymentPage(drawing: Drawing): Promise<PageBuilder> {
	const page = drawing.page(DEPLOYMENT_PAGE);
	page.deferFields = true;
	const networks = await drawNetworks(drawing, page);
	const hosts = drawHosts(page);
	drawLinks(page, hosts);
	drawSidebar(page);
	await drawing.place(page);
	await drawing.edit(
		networks.map((shapeId, index) => ({
			type: 'set-container-members' as const,
			pageId: page.pageId,
			shapeId,
			memberIds: HOSTS.filter((row) => row[2] === index).map((row) =>
				String(hosts[row[0]]?.id ?? 0),
			),
		})),
		'network membership',
	);
	return page;
}
