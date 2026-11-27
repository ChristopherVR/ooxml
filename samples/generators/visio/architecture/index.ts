/** Sample 2: the Nimbuswright platform behind the Aerie weather stations. */
import type { SampleDefinition } from '../../lib/types';
import { Drawing } from '../lib/drawing';
import { PAGE, TITLE_BLOCK, drawTitleBlock } from '../lib/furniture';
import { INK, SHARED_STYLES, swatchStyles } from '../lib/palette';
import { sampleTheme } from '../lib/theme';
import { drawDeploymentPage } from './deployment';
import { DEPLOYMENT_PAGE, LOGICAL_PAGE, drawLogicalPage } from './logical';
import { ARCHITECTURE_MASTERS, ARCHITECTURE_STYLES } from './stencil';

const TITLE = 'Aerie platform architecture';
const DESCRIPTION =
	'The Nimbuswright platform behind the Aerie home weather stations: a logical view by tier and the network it is deployed on.';

async function generate(): Promise<Uint8Array> {
	const drawing = await Drawing.create({
		width: PAGE.w,
		height: PAGE.h,
		ink: INK,
		styles: [...SHARED_STYLES, ...swatchStyles(), ...ARCHITECTURE_STYLES],
		masters: ARCHITECTURE_MASTERS,
		pages: [
			{ name: LOGICAL_PAGE, backPage: TITLE_BLOCK },
			{ name: DEPLOYMENT_PAGE, backPage: TITLE_BLOCK },
			{ name: TITLE_BLOCK, background: true },
		],
		title: TITLE,
		subject: 'Nimbuswright platform engineering',
		description: DESCRIPTION,
		keywords: ['architecture', 'network', 'containers', 'layers'],
		theme: sampleTheme('Nimbuswright'),
	});
	await drawTitleBlock(drawing, {
		organisation: 'Nimbuswright',
		title: TITLE,
		document: 'Nimbuswright · Platform engineering · Architecture record 12',
		revision: '15 Jan 2026',
		accent: 'Swatch sea',
		sheets: 2,
	});
	const logical = await drawLogicalPage(drawing);
	const deployment = await drawDeploymentPage(drawing);
	// Last: once a text field reads Prop.IP, the library refuses to set that row anywhere.
	await drawing.applyFields(logical);
	await drawing.applyFields(deployment);
	return drawing.finish();
}

export const architecture: SampleDefinition = {
	id: 'platform-architecture',
	product: 'visio',
	fileName: 'platform-architecture.vsdx',
	title: 'Platform architecture',
	description: DESCRIPTION,
	features: [
		'Containers with members',
		'Group masters and glued connectors',
		'Layers per tier and per kind of flow',
		'Shape data linked to an inventory recordset',
		'Callouts, text fields and a title block',
	],
	generate,
};
