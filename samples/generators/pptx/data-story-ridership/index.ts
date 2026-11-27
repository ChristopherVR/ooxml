/**
 * "How Port Averly moved": Port Averly Transit's 2025 ridership review, a data story told
 * through native charts on a light editorial theme.
 */
import type { SampleDefinition } from '../../lib/types';
import { Deck } from '../kit/deck';
import { repairPackage } from './fixes';
import { evidenceSlides } from './slides-evidence';
import { headlineSlides } from './slides-headline';
import { movedSlides } from './slides-moved';
import { nextSlides } from './slides-next';
import { whySlides } from './slides-why';
import { TEMPLATE } from './template';

async function generate(): Promise<Uint8Array> {
	const deck = await Deck.open(TEMPLATE);
	headlineSlides(deck);
	movedSlides(deck);
	whySlides(deck);
	evidenceSlides(deck);
	nextSlides(deck);
	return repairPackage(await deck.save());
}

export const sample: SampleDefinition = {
	id: 'data-story-ridership',
	product: 'pptx',
	fileName: 'data-story-ridership.pptx',
	title: 'How Port Averly moved',
	description:
		"Port Averly Transit's 2025 ridership review, told as a data story: six native charts with annotations, a network map and a plan for next year, on a light editorial theme.",
	features: [
		'Custom theme, slide master and six layouts',
		'Native column, line, stacked bar, doughnut, combo and scatter charts with embedded workbooks',
		'Morph, fade, push, wipe, split and cover transitions',
		'Entrance, emphasis, exit and motion-path animations',
		'Text built paragraph by paragraph',
		'Animations triggered by clicking a shape',
		'SmartArt, a formatted table and a generated network map',
		'Gradient fills, soft shadows and custom geometry',
		'Speaker notes on every slide and named sections',
	],
	generate,
};
