/**
 * The PowerPoint samples. Each deck is built with the core's pptx SDK on a hand-written template
 * (see `kit/`), then saved by the core's own `PptxHandler`.
 */
import type { SampleDefinition } from '../lib/types';
import { sample as dataStoryRidership } from './data-story-ridership/index';
import { sample as productLaunchKeynote } from './product-launch-keynote/index';
import { sample as studioPortfolio } from './studio-portfolio/index';

export const samples: SampleDefinition[] = [
	productLaunchKeynote,
	dataStoryRidership,
	studioPortfolio,
];
