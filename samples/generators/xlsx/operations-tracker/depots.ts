/** The Depots sheet: a lookup table with site links, protected except for its staffing inputs. */
import type { Book, Look } from '../lib/book';
import { DEPOTS } from './data';
import { FMT, FOOTER, INPUT, P, SHEETS, title } from './style';

/** The mail address of a person at the company, on the reserved test domain. */
const mail = (name: string): string =>
	`${name.toLowerCase().replaceAll(' ', '.')}@valeorrin.example.invalid`;

export function buildDepots(book: Book): void {
	const sheet = book.sheet(SHEETS.depots);
	title(
		sheet,
		'Depots',
		'The eight sites and who runs them. The sheet is protected without a password; the shaded staffing cells stay editable.',
		P.teal,
	);
	sheet.page(FOOTER, { tall: 1 });
	sheet.ws.view.showGridLines = true;
	sheet.widths({ A: 2, B: 25, C: 18, D: 17, E: 21, F: 13, G: 15, H: 13, I: 52, J: 43 });
	const head: Look = { bold: true, color: P.white, v: 'center' };
	const names = [
		'Depot',
		'City',
		'Region',
		'Manager',
		'Service bays',
		'Technicians on shift',
		'Jobs per bay',
		'Site page',
		'Manager email',
	];
	names.forEach((name, i) =>
		sheet.set(`${String.fromCharCode(66 + i)}5`, name, {
			...head,
			h: i >= 4 && i <= 6 ? 'right' : 'left',
			wrap: true,
			// The link columns follow a column of figures: an indent keeps them apart.
			...(i >= 7 ? { indent: 1 } : {}),
		}),
	);
	sheet.heights({ 5: 30 });
	const link: Look = { color: P.blue, underline: true, indent: 1 };
	const perBay =
		'COUNTIFS(WO_Depot,Depots[[#This Row],[Depot]],WO_Status,"<>Closed")/Depots[[#This Row],[Service bays]]';
	DEPOTS.forEach((depot, n) => {
		const r = 6 + n;
		sheet
			.set(`B${r}`, depot.name, { bold: true })
			.set(`C${r}`, depot.city)
			.set(`D${r}`, depot.region);
		sheet.set(`E${r}`, depot.manager);
		sheet.set(`F${r}`, depot.bays, { ...INPUT, bold: false, fmt: FMT.count });
		sheet.set(`G${r}`, depot.technicians.length + (depot.bays > 5 ? 2 : 1), {
			...INPUT,
			bold: false,
			fmt: FMT.count,
		});
		sheet.formula(`H${r}`, perBay, { fmt: '0.0' });
		const url = `https://depots.valeorrin.example.invalid/${depot.slug}`;
		sheet.set(`I${r}`, url.replace('https://', ''), link);
		sheet.linkOut(`I${r}`, url, `Site page of ${depot.name}`);
		sheet.set(`J${r}`, mail(depot.manager), link);
		sheet.linkOut(`J${r}`, `mailto:${mail(depot.manager)}`, `Write to ${depot.manager}`);
	});
	const last = 5 + DEPOTS.length;
	sheet.ws.tables.push({
		id: 1,
		name: 'Depots',
		displayName: 'Depots',
		range: { start: { row: 4, col: 1 }, end: { row: last - 1, col: 9 } },
		headerRow: true,
		totalsRow: false,
		columns: names.map((name) => ({
			name,
			...(name === 'Jobs per bay' ? { calculatedColumnFormula: perBay } : {}),
		})),
		styleName: 'TableStyleMedium2',
		showRowStripes: true,
	});
	sheet.validate(`F6:G${last}`, {
		type: 'whole',
		operator: 'between',
		formula1: '0',
		formula2: '40',
		allowBlank: false,
		showInputMessage: true,
		promptTitle: 'Staffing',
		prompt: 'A whole number from 0 to 40. These are the only cells you can edit on this sheet.',
		showErrorMessage: true,
		errorStyle: 'stop',
		errorTitle: 'Not a valid count',
		error: 'Enter a whole number between 0 and 40.',
	});
	sheet.conditional(`H6:H${last}`, {
		type: 'iconSet',
		iconSet: '3TrafficLights1',
		reverse: true,
		thresholds: [
			{ type: 'percent', value: '0' },
			{ type: 'num', value: '2.5' },
			{ type: 'num', value: '4' },
		],
	});
	sheet.set(
		`B${last + 2}`,
		'Protected without a password: Review > Unprotect Sheet removes the protection. Shaded cells are unlocked.',
		{ italic: true, color: P.muted, size: 9 },
	);
	sheet.ws.protection = { sheet: true, allow: ['formatColumns', 'sort', 'autoFilter'] };
}
