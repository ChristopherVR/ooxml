/** Sample 3: the Saltmarsh Institute's organisation chart and its pilot project's timeline. */
import type { SampleDefinition } from '../../lib/types';
import { Drawing } from '../lib/drawing';
import { PAGE, TITLE_BLOCK, drawTitleBlock } from '../lib/furniture';
import { INK, SHARED_STYLES, swatchStyles } from '../lib/palette';
import { sampleTheme } from '../lib/theme';
import { ORGANISATION_PAGE, TIMELINE_PAGE, drawOrganisationPage } from './organisation';
import { ORG_MASTERS, ORG_STYLES } from './stencil';
import { drawTimelinePage } from './timeline';

const TITLE = 'Institute organisation';
const DESCRIPTION =
	'Who reports to whom at the Saltmarsh Institute for Coastal Engineering, with the timeline of its Greywharf tidal energy pilot on a second page.';

async function generate(): Promise<Uint8Array> {
	const drawing = await Drawing.create({
		width: PAGE.w,
		height: PAGE.h,
		ink: INK,
		styles: [...SHARED_STYLES, ...swatchStyles(), ...ORG_STYLES],
		masters: ORG_MASTERS,
		pages: [
			{ name: ORGANISATION_PAGE, backPage: TITLE_BLOCK },
			{ name: TIMELINE_PAGE, backPage: TITLE_BLOCK },
			{ name: TITLE_BLOCK, background: true },
		],
		title: TITLE,
		subject: 'Saltmarsh Institute for Coastal Engineering',
		description: DESCRIPTION,
		keywords: ['organisation chart', 'timeline', 'data linking', 'shape data'],
		theme: sampleTheme('Saltmarsh'),
	});
	await drawTitleBlock(drawing, {
		organisation: 'Saltmarsh Institute for Coastal Engineering',
		title: TITLE,
		document: 'Saltmarsh Institute for Coastal Engineering \u00b7 Staff handbook',
		revision: '15 Jan 2026',
		accent: 'Swatch amber',
		sheets: 2,
	});
	const organisation = await drawOrganisationPage(drawing);
	const timeline = await drawTimelinePage(drawing);
	// Last: once a text field reads a Shape Data row, the library refuses to set that row.
	await drawing.applyFields(organisation);
	await drawing.applyFields(timeline);
	return drawing.finish();
}

export const orgChart: SampleDefinition = {
	id: 'institute-org-chart',
	product: 'visio',
	fileName: 'institute-org-chart.vsdx',
	title: 'Institute organisation chart',
	description: DESCRIPTION,
	features: [
		'Organisation tree with glued connectors',
		'Group master with department colours',
		'Cards linked to a staff directory recordset',
		'Project timeline with milestones and dependencies',
		'Date and text fields from shape data',
	],
	generate,
};
