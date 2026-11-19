/**
 * What every {@link PptxAiBridge} must guarantee to the proposal store and the
 * write tools: a read made right after a write returns that write, with no
 * render, tick or await in between. "Apply all" replays several proposals in one
 * tick, each on the deck the previous one left, so a bridge that serves the last
 * rendered state silently drops all but the final write to each field. Writes
 * also keep the deck load, or every staged proposal would be dropped.
 *
 * @module ai/bridge-contract
 */

import type {
	PptxData,
	PptxTheme,
	PptxThemeColorScheme,
	PptxThemeFontScheme,
} from 'ooxml-core/pptx';

import type { PptxAiBridge } from './bridge';
import { readLiveDeck } from './proposals';

const ACCENT_1 = 'A11111';
const ACCENT_2 = 'A22222';
const HEADING_FONT = 'Bridge Contract Heading';
const TITLE = 'Bridge contract title';
const AUTHOR = 'Bridge contract author';
const X = 1111;
const Y = 2222;

function colorOf(theme: PptxTheme | undefined, key: 'accent1' | 'accent2'): string {
	return String(theme?.colorScheme?.[key] ?? '')
		.replace(/^#/u, '')
		.toUpperCase();
}

function firstElement(deck: Pick<PptxData, 'slides'>) {
	return deck.slides[0]?.elements[0];
}

/**
 * Drive `bridge` through consecutive slide, theme and deck writes in a single
 * tick and report every read that did not see the write before it. Returns an
 * empty list for a conforming bridge. The deck needs at least one element on
 * its first slide, and is left edited: run this on a throwaway deck.
 */
export function checkPptxAiBridgeReadsItsWrites(bridge: PptxAiBridge): string[] {
	const failures: string[] = [];
	const check = (what: string, passed: boolean): void => {
		if (!passed) {
			failures.push(what);
		}
	};

	const load = bridge.getDeckLoad();
	const moveFirstElement = (patch: { x?: number; y?: number }, label: string): void => {
		bridge.applySlidesUpdate((slides) => {
			const element = slides[0]?.elements[0];
			if (element) {
				Object.assign(element, patch);
			}
			return slides;
		}, label);
	};
	moveFirstElement({ x: X }, 'Bridge contract: first slide write');
	check('getSlides after a slide write', firstElement({ slides: bridge.getSlides() })?.x === X);
	moveFirstElement({ y: Y }, 'Bridge contract: second slide write');
	const moved = firstElement({ slides: bridge.getSlides() });
	check('a second slide write keeps the first', moved?.x === X && moved?.y === Y);
	check('the whole-deck read after a slide write', firstElement(readLiveDeck(bridge))?.y === Y);

	const scheme = (accent: 'accent1' | 'accent2', value: string): Partial<PptxTheme> => ({
		colorScheme: {
			...bridge.getTheme()?.colorScheme,
			[accent]: `#${value}`,
		} as PptxThemeColorScheme,
	});
	bridge.applyTheme(scheme('accent1', ACCENT_1));
	check('getTheme after a colour write', colorOf(bridge.getTheme(), 'accent1') === ACCENT_1);
	bridge.applyTheme(scheme('accent2', ACCENT_2));
	check(
		'a second colour write keeps the first',
		colorOf(bridge.getTheme(), 'accent1') === ACCENT_1 &&
			colorOf(bridge.getTheme(), 'accent2') === ACCENT_2,
	);
	const fontScheme = bridge.getTheme()?.fontScheme;
	bridge.applyTheme({
		fontScheme: {
			...fontScheme,
			majorFont: { ...fontScheme?.majorFont, latin: HEADING_FONT },
		} as PptxThemeFontScheme,
	});
	check(
		'getTheme after a font write',
		bridge.getTheme()?.fontScheme?.majorFont?.latin === HEADING_FONT,
	);
	check('a font write keeps the colours', colorOf(bridge.getTheme(), 'accent2') === ACCENT_2);

	const setCore = (patch: { title?: string; creator?: string }): void => {
		bridge.applyDeckData?.((deck) => {
			deck.coreProperties = { ...deck.coreProperties, ...patch };
			return deck;
		}, 'Bridge contract: properties');
	};
	setCore({ title: TITLE });
	check(
		'getDeckData after a property write',
		bridge.getDeckData?.()?.coreProperties?.title === TITLE,
	);
	setCore({ creator: AUTHOR });
	const core = bridge.getDeckData?.()?.coreProperties;
	check(
		'a second property write keeps the first',
		core?.title === TITLE && core?.creator === AUTHOR,
	);
	check('the deck load after writes', bridge.getDeckLoad() === load);
	return failures;
}
