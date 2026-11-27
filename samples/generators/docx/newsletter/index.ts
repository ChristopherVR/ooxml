/**
 * The Garden Gate: a two-page newsletter. One continuous run of sections switches between one,
 * two and three columns; the masthead, the pull quote and the pictures float at fixed positions.
 */
import type { HeaderFooterContent } from '../../../../src/core/docx/index';
import type { SampleDefinition } from '../../lib/types';
import { Doc, t } from '../kit/builder';
import { field } from '../kit/fields';
import { finish } from '../kit/finish';
import { numberingXml, settingsXml } from '../kit/parts';
import { save, section } from '../kit/save';
import { stylesXml } from '../kit/styles';
import { tw } from '../kit/tables';
import {
	diary,
	GARDEN_NAME,
	GUTTER,
	ISSUE,
	lead,
	mastheadBlock,
	NAME,
	NOTES,
	sowing,
	TOP,
} from './content';
import { LISTS, MARGIN, STYLE_SHEET, THEME } from './styles';

function footer(doc: Doc, part: string): HeaderFooterContent {
	return {
		partName: `word/footer${part}.xml`,
		blocks: [
			doc.p('Footer', [
				t(GARDEN_NAME, { bold: true }),
				t('  ·  14 Foxglove Rise, Larkspur Heights  ·  open daily, dawn to dusk'),
				t('\tPage '),
				...field('PAGE', '1'),
				t(' of '),
				...field('NUMPAGES', '2'),
			]),
		],
	};
}

async function generate(): Promise<Uint8Array> {
	const doc = new Doc();
	const mast = mastheadBlock(doc);
	const { head, story } = lead(doc);
	const diaryBlocks = diary(doc);
	const notesHeading = doc.p('Heading2', 'Plot notes');
	const notes = NOTES.flatMap(([title, body]) => [doc.p('Heading3', title), doc.p('Normal', body)]);
	const sowHeading = doc.p('Heading2', 'What to sow now');
	const sow = sowing(doc);
	const margins: [number, number, number, number] = [
		Math.round((TOP / 2.54) * 1440),
		MARGIN,
		1247,
		MARGIN,
	];
	const columns = (count: number, separator = false) => ({
		count,
		equalWidth: true,
		spacingTwips: tw(GUTTER),
		...(separator && { separator: true }),
	});
	const base = { margins, headerDistance: 624, footerDistance: 510, type: 'continuous' as const };
	const sections = [
		section({
			...base,
			type: 'nextPage',
			endsAt: head.at(-1)!,
			titlePage: true,
			headers: {
				default: {
					partName: 'word/header1.xml',
					blocks: [doc.p('Header', [t(NAME), t(`\t${ISSUE}`)])],
				},
				first: { partName: 'word/header2.xml', blocks: [doc.p('Normal', '')] },
			},
			footers: { default: footer(doc, '1'), first: footer(doc, '2') },
		}),
		section({ ...base, endsAt: story.at(-1)!, columns: columns(2) }),
		section({ ...base, endsAt: notesHeading }),
		section({ ...base, endsAt: notes.at(-1)!, columns: columns(3, true) }),
		section({ ...base, endsAt: sowHeading }),
		section({ ...base, endsAt: sow.at(-1)!, columns: columns(2) }),
	];
	const bytes = await save(doc, {
		blocks: [mast, ...head, ...story, ...diaryBlocks, notesHeading, ...notes, sowHeading, ...sow],
		sections,
		theme: THEME,
	});
	return finish(bytes, {
		raw: doc.raw,
		styles: stylesXml(STYLE_SHEET),
		numbering: numberingXml(LISTS),
		settings: settingsXml({ language: 'en-GB' }),
		core: {
			title: `${NAME}: ${ISSUE.replace(/\s+·\s+/u, ', ').toLowerCase()}`,
			subject: `Newsletter of the ${GARDEN_NAME}`,
			keywords: 'community garden; newsletter; Larkspur Heights',
			description: 'A sample newsletter with fictitious content, generated with ooxml-core.',
			category: 'Newsletter',
			language: 'en-GB',
		},
		app: { application: 'ooxml-core', company: GARDEN_NAME },
		patch: { 'word/document.xml': roundWrap },
	});
}

/**
 * The core writes a tight wrap with a rectangular outline; the sunflower is round, so its wrap
 * polygon is replaced with a sixteen-sided one (coordinates are in 21600ths of the picture).
 */
function roundWrap(xml: string): string {
	const points = Array.from({ length: 17 }, (_, index) => {
		const angle = (index / 16) * Math.PI * 2 - Math.PI / 2;
		const x = Math.round(10800 + 10800 * Math.cos(angle));
		const y = Math.round(10800 + 10800 * Math.sin(angle));
		return `<wp:${index === 0 ? 'start' : 'lineTo'} x="${x}" y="${y}"/>`;
	});
	const polygon = `<wp:wrapPolygon edited="1">${points.join('')}</wp:wrapPolygon>`;
	const patched = xml.replace(
		/<wp:wrapPolygon[^>]*>(?:(?!<\/wp:wrapPolygon>).)*<\/wp:wrapPolygon>/u,
		polygon,
	);
	if (patched === xml) throw new Error('The tight wrap has no polygon to replace');
	return patched;
}

export const communityNewsletter: SampleDefinition = {
	id: 'community-newsletter',
	product: 'docx',
	fileName: 'community-newsletter.docx',
	title: 'Community newsletter',
	description:
		'A designed two-page newsletter for a community garden: a masthead, columns that change from one to two to three, a drop cap, a pull quote, wrapped pictures, a sidebar and an events calendar, all coloured from a custom theme.',
	features: [
		'Custom theme: colours and fonts used through theme references',
		'Continuous section breaks switching between one, two and three columns',
		'Masthead and pull quote in DrawingML text boxes',
		'Pictures with square and tight wrapping at anchored positions',
		'Drop cap',
		'Sidebar with shading and borders',
		'Events calendar table with a banded table style',
		'Designed header and footer with page numbers',
	],
	generate,
};
