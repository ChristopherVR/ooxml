/** The invented depots, fleet and work-order log of Vale Orrin Logistics, from a fixed seed. */
import { PEOPLE } from '../../lib/fiction';
import { round, seeded } from '../../lib/random';
import { AS_OF } from './style';

export interface Depot {
	name: string;
	slug: string;
	city: string;
	region: string;
	manager: string;
	bays: number;
	technicians: string[];
	/** Relative share of the work. */
	weight: number;
}

const depot = (
	name: string,
	city: string,
	region: string,
	bays: number,
	weight: number,
	index: number,
): Depot => ({
	name,
	slug: name.toLowerCase().replaceAll(' ', '-'),
	city,
	region,
	manager: PEOPLE[index]!,
	bays,
	technicians: [PEOPLE[8 + 2 * index]!, PEOPLE[9 + 2 * index]!],
	weight,
});

export const DEPOTS: Depot[] = [
	depot('Port Averly Hub', 'Port Averly', 'Northwind Coast', 8, 22, 0),
	depot('Marrow Bay Terminal', 'Marrow Bay', 'Northwind Coast', 5, 13, 1),
	depot('Greywharf Quay', 'Greywharf', 'Northwind Coast', 4, 10, 2),
	depot('Kessington Crossdock', 'Kessington', 'Central Plains', 6, 15, 3),
	depot('Elden Ford Yard', 'Elden Ford', 'Central Plains', 4, 9, 4),
	depot('Stonehaven Mills Depot', 'Stonehaven Mills', 'Highland', 3, 8, 5),
	depot('Brackwater Cold Store', 'Brackwater', 'Lakeshore', 4, 11, 6),
	depot('Vale Orrin Central', 'Vale Orrin', 'Southern Delta', 6, 12, 7),
];

export const PRIORITIES = ['Critical', 'High', 'Normal', 'Low'] as const;
export const STATUSES = ['Open', 'In progress', 'Waiting for parts', 'Closed'] as const;
export const CATEGORIES = [
	'Preventive service',
	'Brakes',
	'Tyres',
	'Electrical',
	'Hydraulics',
	'Refrigeration',
	'Bodywork',
	'Engine and driveline',
	'Safety inspection',
] as const;

interface AssetKind {
	label: string;
	prefix: string;
	fleet: number;
	weight: number;
	categories: number[];
}

const KINDS: AssetKind[] = [
	{
		label: 'Tractor unit',
		prefix: 'T',
		fleet: 140,
		weight: 26,
		categories: [0, 0, 1, 2, 3, 7, 7, 8],
	},
	{
		label: 'Rigid truck',
		prefix: 'V',
		fleet: 90,
		weight: 16,
		categories: [0, 0, 1, 2, 3, 6, 7, 8],
	},
	{
		label: 'Reefer trailer',
		prefix: 'R',
		fleet: 110,
		weight: 15,
		categories: [5, 5, 5, 1, 2, 6, 8],
	},
	{ label: 'Curtain trailer', prefix: 'C', fleet: 160, weight: 12, categories: [1, 2, 2, 6, 6, 8] },
	{ label: 'Forklift', prefix: 'F', fleet: 75, weight: 17, categories: [0, 2, 3, 4, 4, 8] },
	{ label: 'Reach truck', prefix: 'K', fleet: 40, weight: 7, categories: [0, 3, 4, 4, 8] },
	{ label: 'Dock leveller', prefix: 'D', fleet: 48, weight: 4, categories: [4, 4, 8, 3] },
	{ label: 'Conveyor line', prefix: 'L', fleet: 18, weight: 3, categories: [3, 3, 0, 8] },
];

export interface WorkOrder {
	id: string;
	opened: number;
	depot: string;
	asset: string;
	category: string;
	priority: (typeof PRIORITIES)[number];
	technician: string;
	status: (typeof STATUSES)[number];
	due: number;
	/** Absent while the work is not closed. */
	closed?: number;
	hours: number;
	cost: number;
}

export const ORDER_COUNT = 4000;
/** 2 January 2025: the first day of the log. */
const FIRST_DAY = 45659;
/** Days allowed from opening to closing, by priority. */
export const SLA = { Critical: 2, High: 3, Normal: 10, Low: 21 } as const;

/** Days past due for work that is not closed, as the Overdue column computes it. */
export const overdueDays = (order: Pick<WorkOrder, 'closed' | 'due'>): number =>
	order.closed === undefined && order.due < AS_OF ? AS_OF - order.due : 0;

function weighted<T>(next: number, items: readonly T[], weight: (item: T) => number): T {
	let at = next * items.reduce((sum, item) => sum + weight(item), 0);
	for (const item of items) {
		at -= weight(item);
		if (at < 0) return item;
	}
	return items[items.length - 1]!;
}

/** Chosen so the log holds each situation the tracker reports, overdue critical work included. */
const SEED = 20260132;

/** The log in the order the sheet stores it: overdue work first, then newest first. */
export function workOrders(seed = SEED): WorkOrder[] {
	const random = seeded(seed);
	const span = AS_OF - 1 - FIRST_DAY;
	const drafts: Omit<WorkOrder, 'id'>[] = [];
	for (let i = 0; i < ORDER_COUNT; i++) {
		const opened = FIRST_DAY + random.int(0, span);
		const home = weighted(random.next(), DEPOTS, (d) => d.weight);
		const kind = weighted(random.next(), KINDS, (k) => k.weight);
		const category = CATEGORIES[random.pick(kind.categories)]!;
		const priority = weighted(random.next(), PRIORITIES, (p) =>
			p === 'Critical' ? 8 : p === 'High' ? 22 : p === 'Normal' ? 50 : 20,
		);
		const due = opened + SLA[priority];
		// Most work closes inside its service level; one job in six runs late.
		const late = random.next() < 0.18;
		const duration = Math.max(
			0,
			Math.round(SLA[priority] * (late ? random.range(1.15, 3.2) : random.range(0.15, 1))),
		);
		// A few recent routine jobs are stuck waiting for a part.
		const stuck =
			priority !== 'Critical' &&
			priority !== 'High' &&
			opened > AS_OF - 150 &&
			random.next() < 0.03;
		const finished = opened + duration;
		const closed = !stuck && finished < AS_OF ? finished : undefined;
		const status =
			closed !== undefined
				? 'Closed'
				: stuck
					? 'Waiting for parts'
					: weighted(random.next(), STATUSES.slice(0, 3), (s) =>
							s === 'Open' ? 45 : s === 'In progress' ? 38 : 17,
						);
		const effort = category === 'Preventive service' ? random.range(1.5, 5) : random.range(0.5, 11);
		const parts =
			random.next() < 0.22
				? 0
				: random.range(18, category === 'Engine and driveline' ? 2600 : 1100);
		drafts.push({
			opened,
			depot: home.name,
			asset: `${kind.label} ${kind.prefix}-${String(random.int(1, kind.fleet)).padStart(3, '0')}`,
			category,
			priority,
			technician: random.pick(home.technicians),
			status,
			due,
			...(closed === undefined ? {} : { closed }),
			// Open work has booked only part of its labour and parts so far.
			hours: round((closed === undefined ? random.range(0, 0.5) : 1) * effort * 4, 0) / 4,
			cost: round((closed === undefined && status === 'Open' ? 0 : 1) * parts, 0),
		});
	}
	return drafts
		.map((order, index) => ({ order, index }))
		.sort((a, b) => a.order.opened - b.order.opened || a.index - b.index)
		.map(({ order }, i) => ({ ...order, id: `WO-${20001 + i}` }))
		.reverse()
		.map((order, index) => ({ order, index }))
		.sort((a, b) => overdueDays(b.order) - overdueDays(a.order) || a.index - b.index)
		.map(({ order }) => order);
}
