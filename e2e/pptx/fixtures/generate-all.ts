/**
 * Every fixture the browser suite generates, in one place: `global-setup.ts`
 * runs it before each Playwright run, and `check-fixtures.ts` runs it in CI to
 * prove the committed decks are what the generators write today.
 *
 * Each generator writes through `writeFixtureDeterministic` (fixed zip dates,
 * order and core-properties stamps) and runs under a `Math.random` seeded from
 * its own name, because the core's save mints random ids for new content (a
 * table's `a16:colId`s). Together they keep a deck whose content has not
 * changed byte-identical on disk across runs.
 */
import { generateBarPictureFillHillFixture } from './generate-bar-picture-fill-hill-fixture';
import { generateBar3DHorizontalFixture } from './generate-bar3d-horizontal-fixture';
import { generateBar3DPictureFillFixture } from './generate-bar3d-picture-fill-fixture';
import { generateFixture as generateBoxCubeTransitionFixture } from './generate-box-cube-transition-fixture';
import { generateChartFixture } from './generate-chart-fixture';
import { generateChartPieBestFitFixture } from './generate-chart-pie-best-fit-fixture';
import { generateChartStylePaletteFixture } from './generate-chart-style-palette-fixture';
import { generateChartTopAxisFixture } from './generate-chart-top-axis-fixture';
import { generateChartUserShapeGroupFixture } from './generate-chart-user-shape-group-fixture';
import { generateFixture as generateCinematicFragmentsFixture } from './generate-cinematic-fragments-fixture';
import { generateCjkLineBreakingFixture } from './generate-cjk-line-breaking-fixture';
import { generateDegenerateShapeFixture } from './generate-degenerate-shape-fixture';
import { generateFixture as generateEffectSoundGalleryFixture } from './generate-effect-sound-gallery-fixture';
import { generateFixture as generateFidelityShowcaseFixture } from './generate-fidelity-showcase-fixture';
import { generateFieldSubstitutionFixture } from './generate-field-substitution-fixture';
import { generateFixture } from './generate-format-painter-fixture';
import { generateLineFillFidelityFixture } from './generate-line-fill-fidelity-fixture';
import { generateLinkedTextBoxFixture } from './generate-linked-textbox-fixture';
import { generateMasterViewsFixture } from './generate-master-views-fixture';
import { generateFixture as generateMediaBookmarkTriggerFixture } from './generate-media-bookmark-trigger-editable-fixture';
import { generateMorphShapeSwapFixture } from './generate-morph-shape-swap-fixture';
import { generateInkFixture, generateOleFixture } from './generate-ole-ink-fixtures';
import { generateParityWave4Fixture } from './generate-parity-wave4-fixture';
import { generatePie3DFixture } from './generate-pie3d-fixture';
import { generatePresetTextInsetsFixture } from './generate-preset-text-insets-fixture';
import { generateRectPathGradientFixture } from './generate-rectpath-gradient-fixture';
import { generateRibbonGalleriesFixture } from './generate-ribbon-galleries-fixture';
import { generateRunProgramFixture } from './generate-run-program-fixture';
import { generateSmartArtBuildFixture } from './generate-smartart-build-fixture';
import { generateFixture as generateTemplateEditingFixture } from './generate-template-editing-fixture';
import {
	generateTemplateGroupFixture,
	generateTemplateMceFixture,
} from './generate-template-group-fixture';
import { generateTextBodyFixture } from './generate-text-body-fixture';
import { generateTextLayoutFixture } from './generate-text-layout-fixture';
import { generateTextWarpFidelityFixture } from './generate-text-warp-fidelity-fixture';
import { generateThemeColorPickerFixture } from './generate-theme-color-picker-fixture';
import { generateFixture as generateTransitionsAnimationsFixture } from './generate-transitions-animations-fixture';
import { generateUnderlineWordsFixture } from './generate-underline-words-fixture';
import { generateUnderlineWordsRubyTabFixture } from './generate-underline-words-ruby-tab-fixture';

/** Every generator, keyed by a stable name that also seeds its randomness. */
const GENERATORS: ReadonlyArray<readonly [string, () => Promise<unknown>]> = [
	['generateFixture', generateFixture],
	['generateRibbonGalleriesFixture', generateRibbonGalleriesFixture],
	['generateChartFixture', generateChartFixture],
	['generateChartPieBestFitFixture', generateChartPieBestFitFixture],
	['generateChartStylePaletteFixture', generateChartStylePaletteFixture],
	['generateChartTopAxisFixture', generateChartTopAxisFixture],
	['generateChartUserShapeGroupFixture', generateChartUserShapeGroupFixture],
	['generateBar3DHorizontalFixture', generateBar3DHorizontalFixture],
	['generateBar3DPictureFillFixture', generateBar3DPictureFillFixture],
	['generateBarPictureFillHillFixture', generateBarPictureFillHillFixture],
	['generatePie3DFixture', generatePie3DFixture],
	['generateFieldSubstitutionFixture', generateFieldSubstitutionFixture],
	['generateDegenerateShapeFixture', generateDegenerateShapeFixture],
	['generateTransitionsAnimationsFixture', generateTransitionsAnimationsFixture],
	['generateEffectSoundGalleryFixture', generateEffectSoundGalleryFixture],
	['generateMediaBookmarkTriggerFixture', generateMediaBookmarkTriggerFixture],
	['generateBoxCubeTransitionFixture', generateBoxCubeTransitionFixture],
	['generateCinematicFragmentsFixture', generateCinematicFragmentsFixture],
	['generateTemplateEditingFixture', generateTemplateEditingFixture],
	['generateTemplateGroupFixture', generateTemplateGroupFixture],
	['generateTemplateMceFixture', generateTemplateMceFixture],
	['generateMasterViewsFixture', generateMasterViewsFixture],
	['generateOleFixture', generateOleFixture],
	['generateInkFixture', generateInkFixture],
	['generateTextLayoutFixture', generateTextLayoutFixture],
	['generateTextBodyFixture', generateTextBodyFixture],
	['generateTextWarpFidelityFixture', generateTextWarpFidelityFixture],
	['generateLinkedTextBoxFixture', generateLinkedTextBoxFixture],
	['generateMorphShapeSwapFixture', generateMorphShapeSwapFixture],
	['generateLineFillFidelityFixture', generateLineFillFidelityFixture],
	['generateParityWave4Fixture', generateParityWave4Fixture],
	['generateCjkLineBreakingFixture', generateCjkLineBreakingFixture],
	['generateUnderlineWordsFixture', generateUnderlineWordsFixture],
	['generateUnderlineWordsRubyTabFixture', generateUnderlineWordsRubyTabFixture],
	['generatePresetTextInsetsFixture', generatePresetTextInsetsFixture],
	['generateSmartArtBuildFixture', generateSmartArtBuildFixture],
	['generateThemeColorPickerFixture', generateThemeColorPickerFixture],
	['generateRectPathGradientFixture', generateRectPathGradientFixture],
	['generateRunProgramFixture', generateRunProgramFixture],
	['generateFidelityShowcaseFixture', generateFidelityShowcaseFixture],
];

/** A small, fast, seedable PRNG (mulberry32) with `Math.random`'s contract. */
function seededRandom(seed: string): () => number {
	let state = 0;
	for (let i = 0; i < seed.length; i++) {
		state = Math.imul(state ^ seed.charCodeAt(i), 0x9e3779b1);
	}
	return () => {
		state = (state + 0x6d2b79f5) | 0;
		let t = Math.imul(state ^ (state >>> 15), 1 | state);
		t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
		return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
	};
}

/** Regenerate every generated fixture in `e2e/pptx/fixtures`. */
export async function generateAllFixtures(): Promise<void> {
	const realRandom = Math.random;
	try {
		for (const [name, generate] of GENERATORS) {
			Math.random = seededRandom(name);
			await generate();
		}
	} finally {
		Math.random = realRandom;
	}
}
