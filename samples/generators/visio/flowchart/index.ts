/** Sample 1: Kestrelline Outfitters' order-to-delivery process as a cross-functional flowchart. */
import type { SampleDefinition } from '../../lib/types';
import { Drawing } from '../lib/drawing';
import { PAGE, TITLE_BLOCK, drawTitleBlock } from '../lib/furniture';
import { INK, SHARED_STYLES, swatchStyles } from '../lib/palette';
import { sampleTheme } from '../lib/theme';
import { MAIN_PAGE, RETURNS_PAGE, drawMainPage } from './main';
import { drawReturnsPage } from './returns';
import { FLOW_MASTERS, FLOW_STYLES } from './stencil';

const TITLE = 'Order to delivery';
const DESCRIPTION =
	'A cross-functional flowchart of how Kestrelline Outfitters takes an order from checkout to the doorstep, with a second page for returns.';

async function generate(): Promise<Uint8Array> {
	const drawing = await Drawing.create({
		width: PAGE.w,
		height: PAGE.h,
		ink: INK,
		styles: [...SHARED_STYLES, ...swatchStyles(), ...FLOW_STYLES],
		masters: FLOW_MASTERS,
		pages: [
			{ name: MAIN_PAGE, backPage: TITLE_BLOCK },
			{ name: RETURNS_PAGE, backPage: TITLE_BLOCK },
			{ name: TITLE_BLOCK, background: true },
		],
		title: TITLE,
		subject: 'Kestrelline Outfitters operations handbook',
		description: DESCRIPTION,
		keywords: ['flowchart', 'swimlanes', 'order to delivery', 'returns'],
		theme: sampleTheme('Kestrelline'),
	});
	await drawTitleBlock(drawing, {
		organisation: 'Kestrelline Outfitters',
		title: TITLE,
		document: 'Kestrelline Outfitters · Operations handbook · Chapter 4',
		revision: '15 Jan 2026',
		accent: 'Swatch teal',
		sheets: 2,
	});
	const main = await drawMainPage(drawing);
	await drawReturnsPage(drawing);
	// Last: once a text field reads Prop.System, the library refuses to set that row anywhere.
	await drawing.applyFields(main);
	return drawing.finish();
}

export const flowchart: SampleDefinition = {
	id: 'order-to-delivery-flowchart',
	product: 'visio',
	fileName: 'order-to-delivery-flowchart.vsdx',
	title: 'Order to delivery flowchart',
	description: DESCRIPTION,
	features: [
		'Swimlanes with phases',
		'Masters and glued connectors',
		'Shape data linked to a recordset',
		'Text fields, layers and hyperlinks',
		'Background page with a title block',
	],
	generate,
};
