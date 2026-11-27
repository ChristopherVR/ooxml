/** The cover and the front matter: summary, document control, contents and the lists. */
import type { Block, Paragraph, TextRun } from '../../../../src/core/docx/index';
import { rich, t, type Runs } from '../kit/builder';
import { equation, pxFromCm } from '../kit/drawing';
import { field, fieldClose, fieldOpen, pageRef } from '../kit/fields';
import { line, table, tw } from '../kit/tables';
import { coverArt } from './art';
import { pageOf, tocMark, type Report } from './context';
import { AUTHORS, COST, LOSSES, n, PALETTE, REPORT, REVIEWERS, STEADY, YEAR } from './data';
import { PAGE_COUNT } from './pages';
import { TEXT_WIDTH } from './styles';

const WORDS = ['no', 'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine', 'ten'];

const NO_LOOK = {
	firstRow: false,
	lastRow: false,
	firstColumn: false,
	lastColumn: false,
	noHBand: true,
	noVBand: true,
} as const;

/** The cover: art bled to the page edges, the title block and the issue details. */
export function cover(report: Report): Block[] {
	const { doc } = report;
	// 21 cm wide, as the art is drawn: the paragraph that anchors it reserves the space under it.
	const art = doc.picture(coverArt(), {
		name: 'cover',
		widthPx: pxFromCm(21),
		heightPx: pxFromCm(11.86),
		alt: 'Layered teal swells under a low amber disc.',
		title: 'Cover art',
		placement: {
			wrap: 'none',
			behindText: true,
			relativeFrom: 'page',
			offsetXPx: 0,
			relativeFromV: 'page',
			offsetYPx: 0,
		},
	});
	const meta = (label: string, value: string): Paragraph =>
		doc.p('CoverMeta', [t(`${label}  `, { bold: true, color: `#${PALETTE.deep}` }), t(value)]);
	return [
		doc.p('Normal', [art], {
			lineSpacingTwips: tw(6000),
			lineSpacingRule: 'exact',
			spacingAfterTwips: tw(0),
		}),
		doc.p('CoverKicker', `Technical report  ·  ${REPORT.number}`),
		doc.p('Title', REPORT.title),
		doc.p('Subtitle', REPORT.subtitle),
		doc.p(
			'Normal',
			rich(
				`Energy yield, availability, power performance and environmental monitoring of a ${YEAR.ratedMw.toFixed(1)} MW tidal array in the Greywharf Narrows, ${REPORT.period}.`,
			),
			{ indentRightTwips: tw(2400), spacingAfterTwips: tw(1500) },
		),
		meta('Prepared by', AUTHORS.map((author) => author.name).join(', ')),
		meta('Issued', `${REPORT.date}, issue 1.0`),
		doc.p('CoverMeta', t(REPORT.institute, { bold: true, color: `#${PALETTE.teal}` }), {
			spacingBeforeTwips: tw(300),
		}),
		doc.p('CoverMeta', 'Tidewrack Lane, Greywharf, Northwind Coast'),
	];
}

/** Four headline figures in a row, each over its label. */
function glance(report: Report): Block {
	const { doc } = report;
	const cells: [string, string][] = [
		[`${n(YEAR.measured)} MWh`, 'net energy exported'],
		[`${YEAR.capacityFactor.toFixed(1)}%`, 'capacity factor'],
		[`${YEAR.availability.toFixed(1)}%`, 'turbine availability'],
		[`€${COST.lcoe}/MWh`, 'levelised cost of energy'],
	];
	const width = Math.floor(TEXT_WIDTH / 4);
	return table(doc, {
		style: 'LayoutGrid',
		grid: [width, width, width, TEXT_WIDTH - 3 * width],
		headerRows: 0,
		look: NO_LOOK,
		cantSplit: true,
		borders: {
			top: line(PALETTE.teal, 1.5),
			bottom: line(PALETTE.rule, 0.5),
		},
		rows: [
			cells.map(([figure, label]) => ({
				fill: `#${PALETTE.foam}`,
				margins: { top: tw(140), bottom: tw(140), left: tw(140), right: tw(80) },
				content: [
					doc.p('TableText', t(figure, { style: 'Figures', fontSize: 15 }), {
						spacingAfterTwips: tw(20),
					}),
					doc.p('TableText', t(label, { color: `#${PALETTE.slate}`, fontSize: 8.5 })),
				],
			})),
		],
	});
}

function documentControl(report: Report): Block {
	const { doc } = report;
	const row = (label: string, value: string | Runs) => [
		{ content: t(label, { bold: true, color: `#${PALETTE.deep}` }) },
		{ content: value },
	];
	return table(doc, {
		style: 'SaltmarshTable',
		textStyle: 'TableText',
		grid: [2400, TEXT_WIDTH - 2400],
		headerRows: 1,
		look: { ...NO_LOOK, firstRow: true },
		cantSplit: true,
		keepTogether: true,
		rows: [
			[{ content: 'Document control', span: 2 }],
			row('Report number', REPORT.number),
			row('Title', `${REPORT.title}: ${REPORT.subtitle.toLowerCase()}`),
			row('Reporting period', REPORT.period),
			row('Issue', `1.0, ${REPORT.date}, issued for review`),
			row(
				'Prepared by',
				AUTHORS.map((author) => `${author.name} (${author.role.toLowerCase()})`).join('; '),
			),
			row('Reviewed by', `${REVIEWERS.leona.author}; ${REVIEWERS.hugo.author}`),
			row('Length', [...field('NUMPAGES', String(PAGE_COUNT)), t(' pages, including the cover')]),
		],
	});
}

export function summary(report: Report): Block[] {
	const share = Math.round((100 * (LOSSES.commissioning + LOSSES.seal)) / YEAR.outage);
	return [
		report.plainHeading('summary', 'Executive summary'),
		report.p(
			`The Greywharf tidal energy pilot is a ${YEAR.ratedMw.toFixed(1)} MW array of four turbines in the Greywharf Narrows, operated by the ${REPORT.institute} as a research installation. This report covers its first full year of operation, ${REPORT.period}.`,
		),
		report.p(
			`In that year the array exported **${n(YEAR.measured)} MWh**, ${YEAR.ratio.toFixed(1)}% of the energy predicted before construction, at a capacity factor of ${YEAR.capacityFactor.toFixed(1)}%. Turbine availability averaged ${YEAR.availability.toFixed(1)}%. Commissioning and a single gearbox seal failure account for ${share}% of the lost turbine-hours; in the ${WORDS[STEADY.months]} months unaffected by either, the array delivered ${STEADY.ratio.toFixed(1)}% of the modelled energy at ${STEADY.availability.toFixed(1)}% availability.`,
		),
		glance(report),
		report.p(
			`Six environmental indicators were monitored under the consent, and all remained within their thresholds. On the year-one cost record the levelised cost of energy is €${COST.lcoe}/MWh, a figure dominated by fixed costs that a larger array would share.`,
			{},
			{ spacingBeforeTwips: tw(240) },
		),
		report.p(
			'The report recommends recalibrating the power curve of the resource model, holding drivetrain spares at the quay, and extending acoustic monitoring through a second season before any expansion is proposed. The complete monthly record is in {table} on page {page}.',
			{ table: report.tab('monthly'), page: report.page('tab', 'monthly') },
			{ spacingAfterTwips: tw(320) },
		),
		documentControl(report),
	];
}

/** Runs that all sit in one hyperlink to `anchor`, as Word writes a contents entry. */
const linkedRuns = (anchor: string, runs: Runs): Runs =>
	runs.map((run): TextRun => ({ ...run, link: { anchor } }));

/**
 * A table of contents field with its entries already in place. The field begins in the first
 * entry and ends in an empty paragraph after the last, the shape Word itself writes.
 */
function contentsField(
	report: Report,
	instr: string,
	entries: { style: string; anchor: string; runs: Runs }[],
): Paragraph[] {
	const { doc } = report;
	const paragraphs = entries.map((entry, index) =>
		doc.p(entry.style, [
			...(index === 0 ? fieldOpen(instr) : []),
			...linkedRuns(entry.anchor, [
				...entry.runs,
				t('\t'),
				...pageRef(entry.anchor, pageOf(entry.anchor)),
			]),
		]),
	);
	return [
		...paragraphs,
		doc.p('Normal', [fieldClose()], {
			spacingAfterTwips: tw(0),
			lineSpacingTwips: tw(40),
			lineSpacingRule: 'exact',
		}),
	];
}

export function contents(report: Report): Block[] {
	const { doc } = report;
	return [
		doc.p('FrontHeading', 'Contents', { pageBreakBefore: true }),
		...contentsField(
			report,
			'TOC \\o "1-3" \\h \\z \\u',
			report.outline.map((entry) => ({
				style: `TOC${entry.level}`,
				anchor: tocMark('sec', entry.key),
				runs: entry.number ? [t(entry.number), t('\t'), t(entry.text)] : [t(entry.text)],
			})),
		),
	];
}

const SYMBOLS: [string, string, string][] = [
	['A', 'Swept area of one rotor', 'm²'],
	['C_p', 'Power coefficient of the rotor', 'dimensionless'],
	['\\mathrm{CF}', 'Capacity factor: net energy over energy at rated power', '%'],
	['E_{\\mathrm{net}}', 'Net energy exported to the grid', 'MWh'],
	['P', 'Power', 'kW'],
	['r', 'Real discount rate', 'per year'],
	['v', 'Depth-averaged flow speed', 'm/s'],
	['\\rho', 'Density of sea water', 'kg/m³'],
	['\\sigma_v', 'Standard deviation of flow speed within a bin', 'm/s'],
];

export function lists(report: Report): Block[] {
	const { doc } = report;
	const list = (label: 'Figure' | 'Table') =>
		contentsField(
			report,
			`TOC \\h \\z \\c "${label}"`,
			report.captions
				.filter((entry) => entry.label === label)
				.map((entry) => ({
					style: 'TableofFigures',
					anchor: tocMark(label === 'Figure' ? 'fig' : 'tab', entry.key),
					runs: [t(`${label} ${entry.number}  ${entry.text}`)],
				})),
		);
	return [
		doc.p('FrontHeading', 'Figures', { pageBreakBefore: true }),
		...list('Figure'),
		doc.p('FrontHeading', 'Tables', { spacingBeforeTwips: tw(480) }),
		...list('Table'),
		doc.p('FrontHeading', 'Symbols', { spacingBeforeTwips: tw(480) }),
		report.dataTable({
			grid: [1500, TEXT_WIDTH - 1500 - 2000, 2000],
			rows: [
				['Symbol', 'Meaning', 'Unit'],
				...SYMBOLS.map(([symbol, meaning, unit]) => [
					// The space keeps the equation inline: alone in a paragraph Word centres it.
					{ content: [equation(symbol), t(' ')] },
					meaning,
					unit,
				]),
			],
		}),
	];
}
