/**
 * Thimblewick Studio's design portfolio: three case studies with generated imagery, built with
 * the pptx SDK on a hand-written template (see `../kit`).
 */
import type { SampleDefinition } from '../../lib/types';
import { Deck } from '../kit/deck';
import { caseSlides } from './slides-cases';
import { closingSlides } from './slides-close';
import { studioSlides } from './slides-studio';
import { TEMPLATE } from './template';

async function generate(): Promise<Uint8Array> {
	const deck = await Deck.open(TEMPLATE);
	studioSlides(deck);
	caseSlides(deck);
	closingSlides(deck);
	return deck.save();
}

export const sample: SampleDefinition = {
	id: 'studio-portfolio',
	product: 'pptx',
	fileName: 'studio-portfolio.pptx',
	title: 'Studio portfolio',
	description:
		'The portfolio of Thimblewick Studio, a small design studio: three case studies (a bakery identity, garden wayfinding and an outfitter’s field guide) told with large generated imagery, quiet typography and Morph transitions that carry each picture from thumbnail to full page.',
	features: [
		'Custom theme, slide master and six layouts in Rubik and Noto Sans',
		'Morph transitions with named shape pairs, plus fade, push, wipe, cover and split',
		'Entrance, emphasis, exit and motion-path animations, with- and after-previous timing',
		'Text built paragraph by paragraph',
		'Interactive sequences triggered by clicking a picture',
		'SmartArt process diagram',
		'Doughnut chart with an embedded workbook',
		'Table with designed cell formatting',
		'Gradient fills, soft shadows and a custom-geometry studio mark',
		'Generated pictures with alternative text',
		'Speaker notes on every slide and named sections',
	],
	generate,
};
