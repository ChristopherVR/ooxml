/**
 * Result shapes of the handler's whole-deck style operations.
 *
 * They live in this import-free leaf (only the model types) so `types.ts`
 * can name them without pulling the runtime class chain into every module
 * that imports the handler contract.
 */
import type { PptxElement, PptxSlide } from '../types';

/** Which slides a background style lands on. */
export type BackgroundStyleScope =
	| { kind: 'all' }
	| { kind: 'slides'; slideIds: readonly string[] };

/** Slides and template elements after a background style was applied. */
export interface BackgroundStyleApplication {
	slides: PptxSlide[];
	templateElementsBySlideId: Record<string, PptxElement[]>;
}

/** Slides and template (master/layout) elements after style references were resolved again. */
export interface StyleReferenceReResolution {
	slides: PptxSlide[];
	templateElementsBySlideId: Record<string, PptxElement[]>;
}
