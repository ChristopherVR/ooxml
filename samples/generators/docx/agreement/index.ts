/**
 * The Quillhaven services agreement: a form to fill in (content controls and two legacy form
 * fields), a schedule of fees with a calculated total, numbered terms and a signature block.
 */
import type { HeaderFooterContent } from '../../../../src/core/docx/index';
import type { SampleDefinition } from '../../lib/types';
import { Doc, t } from '../kit/builder';
import { pxFromCm } from '../kit/drawing';
import { field } from '../kit/fields';
import { finish } from '../kit/finish';
import { numberingXml, settingsXml } from '../kit/parts';
import { save, section } from '../kit/save';
import { stylesXml } from '../kit/styles';
import { halfPoints, tw } from '../kit/tables';
import { mark } from './art';
import { Controls, partHeading } from './controls';
import { fees } from './fees';
import { clientDetails, projectBrief } from './form';
import { LISTS, MARGIN, STYLE_SHEET, THEME } from './styles';
import { signatures, terms } from './terms';

const CO_OP = 'Quillhaven Design Co-operative';
const PROPERTIES = {
	number: 'QDC-2026-0142',
	manager: 'Amara Diallo',
	version: '3.2',
} as const;
/** The date the DATE field shows: the day the sample claims to have been saved. */
const TODAY = '15 January 2026';

/** The row items of the fee schedule become one repeating section (`w15:repeatingSection`). */
function repeatingSection(xml: string): string {
	const item =
		'<w:sdt><w:sdtPr[^>]*>(?:(?!</w:sdtPr>).)*<w15:repeatingSectionItem/></w:sdtPr><w:sdtContent><w:tr>(?:(?!</w:tr>).)*</w:tr></w:sdtContent></w:sdt>';
	const patched = xml.replace(
		new RegExp(`(?:${item})+`, 'u'),
		(items) =>
			`<w:sdt><w:sdtPr><w:alias w:val="Fee lines"/><w:tag w:val="fees.lines"/><w:id w:val="409999"/><w15:repeatingSection/></w:sdtPr><w:sdtContent>${items}</w:sdtContent></w:sdt>`,
	);
	if (patched === xml) throw new Error('The fee schedule has no repeating items');
	return patched;
}

async function generate(): Promise<Uint8Array> {
	const doc = new Doc();
	const controls = new Controls();
	const logo = doc.picture(mark(), {
		name: 'quillhaven-mark',
		widthPx: pxFromCm(0.62),
		heightPx: pxFromCm(0.62),
		alt: 'Quillhaven mark: a quill on a disc.',
		title: 'Quillhaven mark',
	});
	const front = [
		doc.p('Title', 'Services agreement and client intake form'),
		doc.p('Subtitle', [
			t('Agreement '),
			...field('DOCPROPERTY "Agreement number"', t(PROPERTIES.number, { bold: true })),
			t('   ·   prepared on '),
			...field('DATE \\@ "d MMMM yyyy"', TODAY),
			t('   ·   terms version '),
			...field('DOCPROPERTY "Terms version"', PROPERTIES.version),
		]),
		doc.p(
			'Normal',
			`Thank you for choosing ${CO_OP}. Please complete Parts A and B, check the fees in Part C, and sign in Part E. Grey text shows where to type; lists, dates and check boxes open when you click them.`,
		),
		...clientDetails(doc, controls),
		...projectBrief(doc, controls),
		...fees(doc, controls),
	];
	const termsHeading = partHeading(doc, 'D', 'Terms and conditions');
	const clauses = terms(doc, controls);
	const termsEnd = doc.p('Normal', '', {
		spacingAfterTwips: tw(0),
		lineSpacingTwips: tw(40),
		lineSpacingRule: 'exact',
	});
	const signing = signatures(doc, controls);

	const header: HeaderFooterContent = {
		partName: 'word/header1.xml',
		blocks: [
			doc.p('Header', [
				{ ...logo, positionHalfPoints: halfPoints(-5) },
				t(`  ${CO_OP}`, { style: 'Wordmark' }),
				t('\tAgreement '),
				...field('DOCPROPERTY "Agreement number"', PROPERTIES.number),
			]),
		],
	};
	const footer: HeaderFooterContent = {
		partName: 'word/footer1.xml',
		blocks: [
			doc.p('Footer', [
				t(`${CO_OP}  ·  3 Inkpot Yard, Quillhaven  ·  a workers’ co-operative`),
				t('\tPage '),
				...field('PAGE', '1'),
				t(' of '),
				...field('NUMPAGES', '3'),
			]),
		],
	};
	const margins: [number, number, number, number] = [1560, MARGIN, 1304, MARGIN];
	const sections = [
		section({
			endsAt: termsHeading,
			margins,
			headerDistance: 624,
			footerDistance: 567,
			headers: { default: header },
			footers: { default: footer },
		}),
		section({
			endsAt: termsEnd,
			type: 'continuous',
			margins,
			headerDistance: 624,
			footerDistance: 567,
			columns: { count: 2, equalWidth: true, spacingTwips: tw(510) },
		}),
		section({
			endsAt: signing.at(-1)!,
			type: 'continuous',
			margins,
			headerDistance: 624,
			footerDistance: 567,
		}),
	];
	const bytes = await save(doc, {
		blocks: [...front, termsHeading, ...clauses, termsEnd, ...signing],
		sections,
		theme: THEME,
	});
	return finish(bytes, {
		raw: doc.raw,
		styles: stylesXml(STYLE_SHEET),
		numbering: numberingXml(LISTS),
		settings: settingsXml({ language: 'en-GB' }),
		core: {
			title: 'Services agreement and client intake form',
			subject: `${CO_OP}: services agreement`,
			keywords: 'agreement; intake form; content controls; Quillhaven',
			description:
				'A sample fillable agreement with fictitious content, generated with ooxml-core.',
			category: 'Form',
			language: 'en-GB',
		},
		app: { application: 'ooxml-core', company: CO_OP, manager: PROPERTIES.manager },
		custom: [
			{ name: 'Agreement number', type: 'lpwstr', value: PROPERTIES.number },
			{ name: 'Account manager', type: 'lpwstr', value: PROPERTIES.manager },
			{ name: 'Terms version', type: 'lpwstr', value: PROPERTIES.version },
			{ name: 'Fees agreed', type: 'bool', value: false },
			{ name: 'Review date', type: 'filetime', value: '2026-07-15T00:00:00Z' },
		],
		patch: { 'word/document.xml': repeatingSection },
	});
}

export const servicesAgreementForm: SampleDefinition = {
	id: 'services-agreement-form',
	product: 'docx',
	fileName: 'services-agreement-form.docx',
	title: 'Services agreement form',
	description:
		'A fillable services agreement and client intake form for a design co-operative: every kind of content control, legacy form fields, a fee schedule with a calculated total, numbered legal clauses and a signature block.',
	features: [
		'Content controls: plain and rich text, drop-down list, combo box, date picker, check boxes, picture, building block gallery',
		'Repeating section for fee lines and a locked group around the terms',
		'Titles, tags and placeholder text on every control',
		'Fields: DATE, DOCPROPERTY, FORMTEXT, FORMCHECKBOX, =SUM(ABOVE), REF, HYPERLINK, PAGE and NUMPAGES',
		'Custom document properties read by fields',
		'Numbered legal clauses with multilevel numbering',
		'Two-column section for the terms',
		'Signature block table',
	],
	generate,
};
