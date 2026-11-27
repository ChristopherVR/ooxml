/**
 * The Greywharf technical report. The body is built twice: the first pass only numbers the
 * sections, so the second can store the right text in references to sections that come later.
 */
import type { Block, HeaderFooterContent, Paragraph } from '../../../../src/core/docx/index';
import type { SampleDefinition } from '../../lib/types';
import { t } from '../kit/builder';
import { field } from '../kit/fields';
import { finish } from '../kit/finish';
import { numberingXml, settingsXml } from '../kit/parts';
import { save, section } from '../kit/save';
import { stylesXml } from '../kit/styles';
import { tw } from '../kit/tables';
import { introduction, methods, site } from './chapters-a';
import { monthlyRecord, results } from './chapters-b';
import { conclusions, discussion, notes, references } from './chapters-c';
import { Report, type KnownSections } from './context';
import { REPORT } from './data';
import { contents, cover, lists, summary } from './front';
import { sourcesXml } from './sources';
import { LISTS, MARGIN, STYLE_SHEET, TEXT_WIDTH, THEME, WIDE_WIDTH } from './styles';

const MARGINS: [number, number, number, number] = [1560, MARGIN, 1418, MARGIN];

/** The running header and footer of the body, with their right tab at the text width given. */
function running(report: Report, width: number, part: string) {
	const { doc } = report;
	const tabs: Partial<Paragraph> =
		width === TEXT_WIDTH
			? {}
			: {
					tabStops: [
						{ posTwips: tw(TEXT_WIDTH), align: 'clear' },
						{ posTwips: tw(width), align: 'right' },
					],
				};
	const header: HeaderFooterContent = {
		partName: `word/header${part}.xml`,
		blocks: [
			doc.p(
				'Header',
				[t(REPORT.title, { caps: true, bold: true }), t('\t'), t(REPORT.subtitle)],
				tabs,
			),
		],
	};
	const footer = (name: string): HeaderFooterContent => ({
		partName: `word/footer${name}.xml`,
		blocks: [
			doc.p(
				'Footer',
				[
					t(`${REPORT.institute}  ·  `),
					...field('DOCPROPERTY "Report number"', REPORT.number),
					t('\t'),
					...field('PAGE', '1', { style: 'PageNumber' }),
				],
				tabs,
			),
		],
	});
	return { header, footer };
}

function build(known: KnownSections) {
	const report = new Report(known);
	const { doc } = report;
	const coverBlocks = cover(report);
	// The body first: the contents and the lists are made from the outline it leaves behind.
	const closing = report.p(
		'No indicator came close enough to its threshold to prompt a review, and the monitoring programme continues unchanged into its second year.',
		{},
		{ spacingBeforeTwips: tw(240) },
	);
	const bodyA: Block[] = [
		...introduction(report),
		...site(report),
		...methods(report),
		...results(report),
		closing,
	];
	const wide = monthlyRecord(report);
	const bodyB: Block[] = [
		...discussion(report),
		...conclusions(report),
		...references(report),
		...notes(report),
	];
	// The summary is listed in the contents ahead of the body, so its outline entry moves up.
	const summaryBlocks = summary(report);
	report.outline.unshift(report.outline.pop()!);
	const frontEnd = doc.p('Normal', '', {
		spacingAfterTwips: tw(0),
		lineSpacingTwips: tw(40),
		lineSpacingRule: 'exact',
	});
	const front: Block[] = [...summaryBlocks, ...contents(report), ...lists(report), frontEnd];

	const portrait = running(report, TEXT_WIDTH, '1');
	const landscape = running(report, WIDE_WIDTH, '2');
	const frontFooter: HeaderFooterContent = {
		partName: 'word/footer5.xml',
		blocks: [doc.p('Footer', field('PAGE', 'i', { style: 'PageNumber' }), { align: 'center' })],
	};
	const blank: HeaderFooterContent = {
		partName: 'word/header3.xml',
		blocks: [doc.p('Normal', '', { spacingAfterTwips: tw(0) })],
	};
	const sections = [
		section({ endsAt: coverBlocks.at(-1)!, margins: MARGINS }),
		section({
			endsAt: frontEnd,
			margins: MARGINS,
			pageNumbering: { start: 1, format: 'lowerRoman' },
			headers: { default: blank },
			footers: { default: frontFooter },
		}),
		section({
			endsAt: closing,
			margins: MARGINS,
			titlePage: true,
			pageNumbering: { start: 1, format: 'decimal' },
			headers: { default: portrait.header, first: blank },
			footers: { default: portrait.footer('1'), first: portrait.footer('3') },
		}),
		section({
			endsAt: wide.at(-1)!,
			landscape: true,
			margins: MARGINS,
			pageNumbering: { format: 'decimal' },
			headers: { default: landscape.header },
			footers: { default: landscape.footer('2') },
		}),
		section({
			endsAt: bodyB.at(-1)!,
			margins: MARGINS,
			pageNumbering: { format: 'decimal' },
			headers: { default: portrait.header },
			footers: { default: portrait.footer('4') },
		}),
	];
	return { report, blocks: [...coverBlocks, ...front, ...bodyA, ...wide, ...bodyB], sections };
}

async function generate(): Promise<Uint8Array> {
	const numbered = build(new Map());
	const known = new Map(numbered.report.outline.map((entry) => [entry.key, entry.number]));
	const { report, blocks, sections } = build(known);
	const bytes = await save(report.doc, {
		blocks,
		sections,
		theme: THEME,
		bibliographyXml: sourcesXml(),
	});
	return finish(bytes, {
		raw: report.doc.raw,
		styles: stylesXml(STYLE_SHEET),
		numbering: numberingXml(LISTS),
		settings: settingsXml({ footnotes: true, endnotes: true, language: 'en-GB', math: true }),
		core: {
			title: `${REPORT.title}: ${REPORT.subtitle.toLowerCase()}`,
			subject: 'Tidal energy pilot performance, November 2024 to October 2025',
			keywords: 'tidal energy; Greywharf; availability; power curve; levelised cost',
			description: 'A sample technical report with fictitious content, generated with ooxml-core.',
			category: 'Technical report',
			language: 'en-GB',
		},
		app: { application: 'ooxml-core', company: REPORT.institute },
		custom: [
			{ name: 'Report number', type: 'lpwstr', value: REPORT.number },
			{ name: 'Reporting period', type: 'lpwstr', value: REPORT.period },
			{ name: 'Issue', type: 'lpwstr', value: '1.0' },
		],
	});
}

export const technicalReport: SampleDefinition = {
	id: 'technical-report',
	product: 'docx',
	fileName: 'technical-report.docx',
	title: 'Technical report',
	description:
		'A long, fully structured report on a tidal energy pilot: numbered headings, a table of contents, lists of figures and tables, captions, cross-references, notes, citations, equations, charts and tracked review.',
	features: [
		'Cover page and five sections (roman and arabic page numbers, a landscape page)',
		'Styles with multilevel heading numbering',
		'Table of contents, list of figures and list of tables',
		'Numbered captions and cross-references (SEQ, REF, PAGEREF)',
		'Footnotes and endnotes',
		'Citations and a bibliography',
		'Display and inline equations',
		'Tables with header rows, banding and merged cells',
		'Native charts with embedded workbooks, and pictures',
		'Headers and footers with a different first page',
		'Comments with a reply and tracked changes',
	],
	generate,
};
