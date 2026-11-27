/** The invented catalogue and order book of Kestrelline Outfitters, generated from a fixed seed. */
import { ymdToSerial } from '../../../../src/core/xlsx/formula/date-serial';
import { PEOPLE, REGIONS } from '../../lib/fiction';
import { seeded } from '../../lib/random';

export interface Product {
	name: string;
	category: string;
	price: number;
	cost: number;
	/** Relative demand. */
	weight: number;
	/** The months (1 to 12) the product sells best in. */
	peak: number[];
}

const SUMMER = [5, 6, 7, 8];
const SHOULDER = [3, 4, 9, 10];
const WINTER = [11, 12, 1, 2];
const p = (
	name: string,
	category: string,
	price: number,
	cost: number,
	weight: number,
	peak: number[],
): Product => ({ name, category, price, cost, weight, peak });

export const PRODUCTS: Product[] = [
	p('Ridgewren 2 tent', 'Tents and shelters', 389, 205, 7, SUMMER),
	p('Skylark Solo tent', 'Tents and shelters', 279, 150, 5, SUMMER),
	p('Tarnhold 4 basecamp tent', 'Tents and shelters', 549, 300, 3, SUMMER),
	p('Windbreak tarp', 'Tents and shelters', 89, 38, 4, SHOULDER),
	p('Featherfall 20 sleeping bag', 'Sleep systems', 259, 128, 6, SHOULDER),
	p('Emberdown 0 sleeping bag', 'Sleep systems', 399, 205, 4, WINTER),
	p('Mossbed sleeping pad', 'Sleep systems', 119, 52, 7, SUMMER),
	p('Driftdown camp pillow', 'Sleep systems', 34, 12, 5, SUMMER),
	p('Wren 28 daypack', 'Packs', 129, 58, 9, SHOULDER),
	p('Corrie 45 trekking pack', 'Packs', 219, 104, 6, SUMMER),
	p('Tarnside 65 expedition pack', 'Packs', 299, 150, 3, SUMMER),
	p('Burnfoot 12 running vest', 'Packs', 99, 44, 4, SHOULDER),
	p('Fellwalker boot', 'Footwear', 229, 112, 9, SHOULDER),
	p('Scree approach shoe', 'Footwear', 159, 76, 6, SUMMER),
	p('Bogtrot trail runner', 'Footwear', 139, 66, 8, SHOULDER),
	p('Heather camp sandal', 'Footwear', 59, 24, 4, SUMMER),
	p('Gale shell jacket', 'Apparel', 319, 146, 8, WINTER),
	p('Hearth fleece midlayer', 'Apparel', 109, 44, 9, WINTER),
	p('Drizzle overtrousers', 'Apparel', 129, 56, 5, WINTER),
	p('Thistle wool base layer', 'Apparel', 79, 31, 8, WINTER),
	p('Kettlecairn stove', 'Camp kitchen', 94, 41, 6, SUMMER),
	p('Coalglow titanium pot', 'Camp kitchen', 64, 26, 5, SUMMER),
	p('Rillwater water filter', 'Camp kitchen', 54, 21, 6, SUMMER),
	p('Mothlight camp lantern', 'Camp kitchen', 44, 17, 5, WINTER),
];

export const CATEGORIES = [...new Set(PRODUCTS.map((product) => product.category))];
export const CHANNELS = ['Store', 'Online', 'Wholesale'] as const;

export interface Territory {
	region: string;
	cities: string[];
	reps: string[];
	manager: string;
	target: number;
	weight: number;
}

export const TERRITORIES: Territory[] = [
	{
		region: REGIONS[0],
		cities: ['Port Averly', 'Marrow Bay', 'Greywharf'],
		reps: [PEOPLE[2], PEOPLE[3]],
		manager: PEOPLE[12],
		target: 92000,
		weight: 28,
	},
	{
		region: REGIONS[1],
		cities: ['Kessington', 'Elden Ford', 'Sunderby'],
		reps: [PEOPLE[4], PEOPLE[5]],
		manager: PEOPLE[13],
		target: 77000,
		weight: 20,
	},
	{
		region: REGIONS[2],
		cities: ['Stonehaven Mills', 'Thistledown'],
		reps: [PEOPLE[6], PEOPLE[7]],
		manager: PEOPLE[14],
		target: 70000,
		weight: 17,
	},
	{
		region: REGIONS[3],
		cities: ['Larkspur Heights', 'Quillhaven', 'Brackwater'],
		reps: [PEOPLE[8], PEOPLE[9]],
		manager: PEOPLE[15],
		target: 72000,
		weight: 20,
	},
	{
		region: REGIONS[4],
		cities: ['Vale Orrin', 'Saltmarsh Ferry'],
		reps: [PEOPLE[10], PEOPLE[11]],
		manager: PEOPLE[16],
		target: 62000,
		weight: 15,
	},
];

export interface Order {
	id: string;
	/** Date serial (1900 system). */
	date: number;
	month: number;
	region: string;
	city: string;
	channel: (typeof CHANNELS)[number];
	rep: string;
	product: Product;
	units: number;
	discount: number;
}

export const YEAR = 2025;
export const ORDER_COUNT = 960;
/** Share of the year's orders placed in each month: a summer peak and a gift season. */
const SEASON = [5, 5, 7, 9, 11, 12, 12, 10, 8, 7, 8, 6];
const DAYS = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

function weighted<T>(next: number, items: readonly T[], weight: (item: T) => number): T {
	let at = next * items.reduce((sum, item) => sum + weight(item), 0);
	for (const item of items) {
		at -= weight(item);
		if (at < 0) return item;
	}
	return items[items.length - 1]!;
}

export function orders(): Order[] {
	const random = seeded(20250115);
	const list: Omit<Order, 'id'>[] = [];
	for (let i = 0; i < ORDER_COUNT; i++) {
		const month = weighted(
			random.next(),
			[1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12],
			(m) => SEASON[m - 1]!,
		);
		const day = random.int(1, DAYS[month - 1]!);
		const territory = weighted(random.next(), TERRITORIES, (t) => t.weight);
		const channel = weighted(random.next(), CHANNELS, (c) =>
			c === 'Store' ? 52 : c === 'Online' ? 38 : 10,
		);
		const product = weighted(
			random.next(),
			PRODUCTS,
			(item) => item.weight * (item.peak.includes(month) ? 2.4 : 1),
		);
		const wholesale = channel === 'Wholesale';
		const units = wholesale ? random.int(4, 14) : random.next() < 0.55 ? 1 : random.int(2, 4);
		const discount = wholesale
			? random.pick([0.25, 0.3, 0.3, 0.35])
			: channel === 'Online'
				? random.pick([0, 0, 0.05, 0.1, 0.15])
				: random.pick([0, 0, 0, 0.1]);
		list.push({
			date: ymdToSerial(YEAR, month, day, false)!,
			month,
			region: territory.region,
			city: random.pick(territory.cities),
			channel,
			rep: random.pick(territory.reps),
			product,
			units,
			discount,
		});
	}
	// The order book is kept in date order, and order numbers follow it.
	return list
		.map((order, index) => ({ order, index }))
		.sort((a, b) => a.order.date - b.order.date || a.index - b.index)
		.map(({ order }, i) => ({ ...order, id: `KO-25-${String(i + 1).padStart(4, '0')}` }));
}

/** What an order earns, as the workbook's own formulas compute it. */
export const revenueOf = (order: Order): number =>
	Math.round(order.units * order.product.price * (1 - order.discount) * 100) / 100;
