/** The Aerie launch keynote: a dark, cinematic product deck built with the pptx SDK. */
import type { SampleDefinition } from '../../lib/types';
import { Deck } from '../kit/deck';
import { removeChartLegend, restyleDiagram } from './patches';
import { launchSlides } from './slides-launch';
import { openingSlides } from './slides-opening';
import { STEP_STYLE, productSlides } from './slides-product';
import { proofSlides } from './slides-proof';
import { template } from './template';

async function generate(): Promise<Uint8Array> {
	const deck = await Deck.open(template());
	openingSlides(deck);
	productSlides(deck);
	proofSlides(deck);
	launchSlides(deck);
	const saved = await removeChartLegend(await deck.save(), 'Pre-orders');
	return restyleDiagram(saved, STEP_STYLE);
}

export const sample: SampleDefinition = {
	id: 'product-launch-keynote',
	product: 'pptx',
	fileName: 'product-launch-keynote.pptx',
	title: 'Product launch keynote',
	description:
		'The launch keynote for Aerie, the home weather station from Nimbuswright: thirteen dark, cinematic slides that Morph a glowing sensor ring from the cover into the product reveal, then prove the claim with a chart, a specification table and a live storm warning.',
	features: [
		'Custom dark theme, slide master and eight layouts',
		'Morph transitions with paired shapes, plus fade, push, wipe, split, cover and zoom',
		'Entrance, emphasis and exit animations with with-previous and after-previous timing',
		'Text built paragraph by paragraph',
		'Motion path animation',
		'Animation triggered by clicking a shape',
		'SmartArt process diagram',
		'Column and doughnut charts with embedded workbooks',
		'Table with designed formatting',
		'Gradient fills, soft shadows and custom geometry',
		'Generated pictures with transparency',
		'Speaker notes on every slide and named sections',
	],
	generate,
};
