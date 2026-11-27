// Sheet label products, from the dimensions their vendors publish: page size, label size, the
// number across and down, and the pitch (label plus gap) in each direction. A margin left out is
// the one that centres the grid on the page, as these sheets are cut. Only products whose
// published dimensions agree with that arithmetic are listed. This module is loaded with the
// rest of the stationery code, on demand.
import { twips, type Twips } from '../../units/units';

export interface LabelSpec {
	vendor: string;
	product: string;
	/** Other product numbers sold with the same layout (inkjet and colour variants). */
	aliases?: readonly string[];
	/** A short description, such as "Address labels". */
	description: string;
	columns: number;
	rows: number;
	pageWidth: Twips;
	pageHeight: Twips;
	left: Twips;
	top: Twips;
	width: Twips;
	height: Twips;
	/** Space between two labels across. */
	gap: Twips;
	/** Space between two labels down. */
	verticalGap: Twips;
}

type Row = [
	product: string,
	description: string,
	columns: number,
	rows: number,
	width: number,
	height: number,
	horizontalPitch: number,
	verticalPitch: number,
	aliases?: readonly string[],
	margins?: { left?: number; top?: number },
];

interface Sheet {
	vendor: string;
	/** Twips per unit of the rows: 1440 for inches, 1440 / 25.4 for millimetres. */
	unit: number;
	page: [width: number, height: number];
	rows: readonly Row[];
}

const INCH = 1440;
const MM = 1440 / 25.4;
const LETTER: Sheet['page'] = [8.5, 11];
const A4: Sheet['page'] = [210, 297];

const SHEETS: readonly Sheet[] = [
	{
		vendor: 'Avery US Letter',
		unit: INCH,
		page: LETTER,
		rows: [
			[
				'5160',
				'Address labels',
				3,
				10,
				2.625,
				1,
				2.75,
				1,
				[
					'5260',
					'5520',
					'5660',
					'5960',
					'8160',
					'8460',
					'8660',
					'18160',
					'18660',
					'28660',
					'48160',
				],
			],
			['5161', 'Address labels', 2, 10, 4, 1, 4.1875, 1, ['8161']],
			['5162', 'Address labels', 2, 7, 4, 4 / 3, 4.1875, 4 / 3, ['8162']],
			['5163', 'Shipping labels', 2, 5, 4, 2, 4.1875, 2, ['5263', '5963', '8163', '8463', '18163']],
			['5164', 'Shipping labels', 2, 3, 4, 10 / 3, 4.1875, 10 / 3, ['5264', '8164']],
			['5165', 'Full-sheet labels', 1, 1, 8.5, 11, 8.5, 11, ['8165']],
			['5126', 'Shipping labels (half sheet)', 1, 2, 8.5, 5.5, 8.5, 5.5, ['8126']],
			['5167', 'Return address labels', 4, 20, 1.75, 0.5, 2.0625, 0.5, ['5267', '8167', '18167']],
			['5195', 'Return address labels', 4, 15, 1.75, 2 / 3, 2.0625, 2 / 3],
			['5168', 'Shipping labels', 2, 2, 3.5, 5, 4, 5, ['8168']],
		],
	},
	{
		vendor: 'Online Labels',
		unit: INCH,
		page: LETTER,
		rows: [
			['OL875', 'Address labels', 3, 10, 2.625, 1, 2.75, 1],
			['OL150', 'Address labels', 2, 7, 4, 4 / 3, 4.1875, 4 / 3],
			['OL125', 'Shipping labels', 2, 5, 4, 2, 4.1875, 2],
		],
	},
	{
		vendor: 'Avery A4',
		unit: MM,
		page: A4,
		rows: [
			['L7159', 'Address labels', 3, 8, 63.5, 33.9, 66.04, 33.9, ['J8159']],
			['L7160', 'Address labels', 3, 7, 63.5, 38.1, 66.04, 38.1, ['J8160']],
			['L7161', 'Address labels', 3, 6, 63.5, 46.6, 66.04, 46.6, ['J8161']],
			['L7162', 'Address labels', 2, 8, 99.1, 33.9, 101.6, 33.9, ['J8162']],
			['L7163', 'Address labels', 2, 7, 99.1, 38.1, 101.6, 38.1, ['J8163']],
			['L7164', 'Labels', 3, 4, 63.5, 72, 66.04, 72],
			['L7165', 'Shipping labels', 2, 4, 99.1, 67.7, 101.6, 67.7, ['J8165']],
			['L7166', 'Shipping labels', 2, 3, 99.1, 93.1, 101.6, 93.1, ['J8166']],
			['L7167', 'Full-sheet labels', 1, 1, 199.6, 289.1, 199.6, 289.1, ['J8167']],
			['L7168', 'Shipping labels', 1, 2, 199.6, 143.5, 199.6, 143.5, ['J8168']],
			['L7169', 'Shipping labels', 2, 2, 99.1, 139, 101.6, 139, ['J8169']],
			['L7173', 'Shipping labels', 2, 5, 99.1, 57, 101.6, 57, ['J8173']],
			['L7651', 'Mini labels', 5, 13, 38.1, 21.2, 40.64, 21.2, ['J8651']],
			['L7654', 'Multipurpose labels', 4, 10, 45.7, 25.4, 48.26, 25.4],
		],
	},
	{
		vendor: 'Avery Zweckform',
		unit: MM,
		page: A4,
		rows: [
			['3422', 'Universal labels', 3, 8, 70, 35, 70, 35],
			['3424', 'Universal labels', 2, 6, 105, 48, 105, 48],
			['3425', 'Universal labels', 2, 5, 105, 57, 105, 57],
			['3427', 'Universal labels', 2, 4, 105, 74, 105, 74],
			['3474', 'Universal labels', 3, 8, 70, 37, 70, 37],
			['3475', 'Universal labels', 3, 8, 70, 36, 70, 36],
			['3478', 'Full-sheet labels', 1, 1, 210, 297, 210, 297],
			['3483', 'Universal labels (A6)', 2, 2, 105, 148, 105, 148],
			['3484', 'Universal labels', 2, 8, 105, 37, 105, 37],
			['3651', 'Universal labels', 4, 10, 52.5, 29.7, 52.5, 29.7],
			['3652', 'Universal labels', 3, 7, 70, 42.3, 70, 42.3],
		],
	},
	{
		vendor: 'HERMA A4',
		unit: MM,
		page: A4,
		rows: [['4464', 'Universal labels', 3, 8, 70, 37, 70, 37]],
	},
];

const round = (value: number) => twips(Math.max(0, Math.round(value)));

function spec(sheet: Sheet, row: Row): LabelSpec {
	const [product, description, columns, rows, width, height, across, down, aliases, margins] = row;
	const { unit } = sheet;
	const [pageWidth, pageHeight] = sheet.page;
	const gridWidth = (columns - 1) * across + width;
	const gridHeight = (rows - 1) * down + height;
	return {
		vendor: sheet.vendor,
		product,
		...(aliases && { aliases }),
		description,
		columns,
		rows,
		pageWidth: round(pageWidth * unit),
		pageHeight: round(pageHeight * unit),
		left: round((margins?.left ?? (pageWidth - gridWidth) / 2) * unit),
		top: round((margins?.top ?? (pageHeight - gridHeight) / 2) * unit),
		width: round(width * unit),
		height: round(height * unit),
		gap: round((across - width) * unit),
		verticalGap: round((down - height) * unit),
	};
}

/** Every listed product, grouped by vendor in the order Word's Label Options lists them. */
export const labelCatalogue: readonly LabelSpec[] = SHEETS.flatMap((sheet) =>
	sheet.rows.map((row) => spec(sheet, row)),
);

const normal = (value: string) => value.toLocaleLowerCase('en').replace(/[\s-]+/gu, '');

/** The product with this number or alias, preferring `vendor` when two vendors share it. */
export function findLabel(product: string, vendor?: string): LabelSpec | undefined {
	const wanted = normal(product);
	const matches = labelCatalogue.filter(
		(item) =>
			normal(item.product) === wanted || item.aliases?.some((alias) => normal(alias) === wanted),
	);
	return matches.find((item) => item.vendor === vendor) ?? matches[0];
}

/**
 * Products whose number, alias or description contains `query` (case, spaces and hyphens
 * ignored), limited to `vendor` when given. An empty query lists the vendor's products.
 */
export function searchLabels(query: string, vendor?: string): LabelSpec[] {
	const term = normal(query);
	return labelCatalogue.filter(
		(item) =>
			(!vendor || item.vendor === vendor) &&
			(!term ||
				normal(item.product).includes(term) ||
				item.aliases?.some((alias) => normal(alias).includes(term)) ||
				normal(item.description).includes(term)),
	);
}

export const labelVendors: readonly string[] = [...new Set(SHEETS.map(({ vendor }) => vendor))];
