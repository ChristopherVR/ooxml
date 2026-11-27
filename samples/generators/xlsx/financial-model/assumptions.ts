/** The Assumptions sheet: every input of the model, validated, and the names formulas use. */
import { SAMPLE_AUTHOR } from '../../lib/types';
import type { Book } from '../lib/book';
import { DRIVERS, SCENARIOS, amount, decimal } from './inputs';
import { SCALARS } from './scalars';
import { FMT, FOOTER, LOOK, P, SHEETS, YEARS, frame, line, section, yearHeader } from './style';

const S = SHEETS.assumptions;

/** Row numbers other sheets read. A driver block's `selected` row follows its three scenarios. */
export const A = {
	newStores: 18,
	likeForLike: 24,
	ticketInflation: 30,
	ingredients: 36,
	labour: 42,
	wholesaleGrowth: 51,
	onlineGrowth: 53,
	extension: 70,
	borrowing: 79,
	repayments: 80,
} as const;

function drivers(book: Book): void {
	const sheet = book.sheet(S);
	DRIVERS.forEach((driver, k) => {
		const top = 14 + 6 * k;
		sheet.set(`B${top}`, driver.title, { bold: true, v: 'center' });
		sheet.set(`C${top}`, driver.unit, LOOK.unit);
		SCENARIOS.forEach((name, i) => {
			const row = top + 1 + i;
			line(sheet, row, {
				label: name,
				cell: (_col, _prev, year) => driver.values[i]![year]!,
				fmt: driver.fmt,
			});
			sheet.validate(`E${row}:I${row}`, driver.rule);
		});
		const selected = top + 4;
		line(sheet, selected, {
			label: 'Selected scenario',
			kind: 'subtotal',
			fmt: driver.fmt,
			cell: (col) => `=INDEX(${col}${top + 1}:${col}${top + 3},MATCH(Scenario,ScenarioList,0))`,
		});
		if (driver.actual !== undefined) {
			sheet.set(`D${selected}`, driver.actual, { named: 'Input', fmt: driver.fmt });
			sheet.validate(`D${selected}`, driver.rule);
		}
		// The three scenario rows fold away under the selected row.
		sheet.groupRows(top + 1, top + 3);
		sheet.conditional(`B${top + 1}:I${top + 3}`, {
			type: 'expression',
			formula: `$B${top + 1}=Scenario`,
			style: {
				font: { bold: true, color: { rgb: `FF${P.ink}` } },
				fill: { type: 'pattern', pattern: 'solid', fgColor: { rgb: 'FFF3D58A' } },
			},
		});
	});
}

export function buildAssumptions(book: Book): void {
	const sheet = book.sheet(S);
	frame(
		sheet,
		'Assumptions',
		'Every input of the model lives here. Shaded cells are inputs; each has a validation rule and an input message.',
		P.honey,
	);
	sheet.page(FOOTER, { portrait: true });
	yearHeader(sheet, 'Assumption', false);
	const head = { bold: true, color: P.white, fill: P.ink, v: 'center', h: 'right' } as const;
	sheet.formula('D5', 'StartYear-1', { ...head, fmt: FMT.actualYear });
	YEARS.forEach((col, i) =>
		sheet.formula(`${col}5`, i === 0 ? 'StartYear' : `${YEARS[i - 1]}5+1`, {
			...head,
			fmt: FMT.forecastYear,
		}),
	);
	sheet.widths({ A: 2, B: 42, C: 13, D: 13, 'E:I': 12.5, J: 2, K: 52 }).freeze(5, 3);

	section(sheet, 7, 'General');
	sheet.set('B8', 'Scenario', LOOK.detail).set('C8', 'list', LOOK.unit);
	sheet.set('D8', 'Base', { named: 'Input', h: 'center', bold: true });
	sheet.validate('D8', {
		type: 'list',
		formula1: 'ScenarioList',
		showDropDown: true,
		allowBlank: false,
		showInputMessage: true,
		promptTitle: 'Scenario',
		prompt: 'Pick the scenario the whole model runs on.',
		showErrorMessage: true,
		errorStyle: 'stop',
		errorTitle: 'Unknown scenario',
		error: 'Choose Base, Expansion or Cautious from the list.',
	});
	sheet.set('K8', 'Switches every "Selected scenario" row below.', LOOK.note);

	section(sheet, 13, 'Scenario drivers');
	sheet.set('K13', 'The row of the scenario in use is highlighted.', LOOK.note);
	drivers(book);

	section(sheet, 44, 'Stores and revenue');
	section(sheet, 55, 'Costs');
	section(sheet, 65, 'Capital expenditure and depreciation');
	section(sheet, 72, 'Working capital');
	section(sheet, 77, 'Financing and distributions');
	for (const scalar of SCALARS) {
		line(sheet, scalar.row, {
			label: scalar.label,
			unit: scalar.unit,
			base: scalar.value,
			fmt: scalar.fmt,
		});
		sheet.validate(`D${scalar.row}`, scalar.rule);
		book.wb.definedNames.push({ name: scalar.name, formula: `${S}!$D$${scalar.row}` });
	}
	const growth = decimal(
		-0.2,
		0.5,
		'Growth rate',
		'Year-on-year growth of this revenue stream.',
		'warning',
	);
	line(sheet, A.wholesaleGrowth, {
		label: 'Wholesale revenue growth',
		unit: '% a year',
		fmt: FMT.percent,
		cell: (_c, _p, i) => [0.06, 0.06, 0.05, 0.05, 0.05][i]!,
	});
	line(sheet, A.onlineGrowth, {
		label: 'Online orders revenue growth',
		unit: '% a year',
		fmt: FMT.percent,
		cell: (_c, _p, i) => [0.18, 0.16, 0.14, 0.12, 0.1][i]!,
	});
	sheet.validate(
		`E${A.wholesaleGrowth}:I${A.wholesaleGrowth} E${A.onlineGrowth}:I${A.onlineGrowth}`,
		growth,
	);
	line(sheet, A.extension, {
		label: 'Central bakery extension',
		unit: '$',
		fmt: FMT.money,
		cell: (_c, _p, i) => [0, 2400000, 600000, 0, 0][i]!,
	});
	sheet.validate(
		`E${A.extension}:I${A.extension}`,
		amount('Central bakery extension', 'Project spending on the extension of the central bakery.'),
	);
	line(sheet, A.borrowing, {
		label: 'New borrowing drawn',
		unit: '$',
		fmt: FMT.money,
		cell: (_c, _p, i) => [0, 1500000, 0, 0, 0][i]!,
	});
	line(sheet, A.repayments, {
		label: 'Scheduled debt repayments',
		unit: '$',
		fmt: FMT.money,
		cell: () => 400000,
	});
	sheet.validate(
		`E${A.borrowing}:I${A.repayments}`,
		amount('Debt schedule', 'Amounts drawn or repaid during the year, in whole dollars.'),
	);
	sheet.set(
		`K${A.borrowing}`,
		'A facility for the central bakery extension is drawn in the second year.',
		LOOK.note,
	);
	sheet.set(
		'K49',
		'The second half trades at 2 minus this index, so the year averages 1.',
		LOOK.note,
	);

	// The same three cases in Excel's Scenario Manager: showing one sets the selector cell.
	sheet.ws.scenarios = {
		current: 0,
		show: 0,
		list: SCENARIOS.map((name) => ({
			name,
			user: SAMPLE_AUTHOR,
			comment: `Runs the model on the ${name} drivers.`,
			cells: [{ ref: 'D8', value: name }],
		})),
	};
	book.wb.definedNames.push(
		{ name: 'Scenario', formula: `${S}!$D$8`, comment: 'The scenario the model runs on' },
		{ name: 'ScenarioList', formula: `${S}!$B$15:$B$17` },
		{ name: 'ForecastYears', formula: `${S}!$E$5:$I$5` },
		// A name that exists only on this sheet: the block of scenario inputs.
		{ name: 'DriverInputs', formula: `${S}!$E$15:$I$41`, localSheet: sheet.index },
	);
}
