/**
 * PowerPoint's Ribbon Display Options: how much of the ribbon stays on screen. The three modes
 * and their order are PowerPoint's own menu (Full-screen mode, Show tabs only, Always show
 * Ribbon). Pure, so every binding draws the same thing from the same decision.
 */

/**
 * - `always`: tabs and commands stay docked above the slide.
 * - `tabs`: only the tab row shows; choosing a tab lays its commands over the slide until a
 *   press elsewhere or Escape.
 * - `fullScreen`: the title bar and the ribbon leave; a slim strip at the top brings the ribbon
 *   back over the slide for one command.
 */
export type RibbonDisplayMode = 'always' | 'tabs' | 'fullScreen';

/** The modes in the order PowerPoint lists them. */
export const RIBBON_DISPLAY_MODES: readonly RibbonDisplayMode[] = ['fullScreen', 'tabs', 'always'];

export const DEFAULT_RIBBON_DISPLAY_MODE: RibbonDisplayMode = 'always';

const LABEL_KEYS: Record<RibbonDisplayMode, string> = {
	fullScreen: 'pptx.ribbon.display.fullScreen',
	tabs: 'pptx.ribbon.display.tabs',
	always: 'pptx.ribbon.display.always',
};

/** Translation key of the menu button itself. */
export const RIBBON_DISPLAY_OPTIONS_LABEL_KEY = 'pptx.ribbon.displayOptions';
/** Translation key of the strip that brings the ribbon back in full-screen mode. */
export const RIBBON_DISPLAY_REVEAL_LABEL_KEY = 'pptx.ribbon.showRibbon';

/** A host value (an option, a stored preference) as a mode; anything else is the default. */
export function normalizeRibbonDisplayMode(value: unknown): RibbonDisplayMode {
	return RIBBON_DISPLAY_MODES.includes(value as RibbonDisplayMode)
		? (value as RibbonDisplayMode)
		: DEFAULT_RIBBON_DISPLAY_MODE;
}

export interface RibbonDisplayMenuItem {
	mode: RibbonDisplayMode;
	labelKey: string;
	checked: boolean;
}

/** The Ribbon Display Options menu, one radio item per mode. */
export function ribbonDisplayMenu(current: RibbonDisplayMode): RibbonDisplayMenuItem[] {
	return RIBBON_DISPLAY_MODES.map((mode) => ({
		mode,
		labelKey: LABEL_KEYS[mode],
		checked: mode === current,
	}));
}

/**
 * Ctrl+F1 and a double click on a tab, as in PowerPoint: they switch between the docked ribbon
 * and the tab row; from full-screen mode they return to the docked ribbon.
 */
export function toggleRibbonCollapsed(mode: RibbonDisplayMode): RibbonDisplayMode {
	return mode === 'always' ? 'tabs' : 'always';
}

export interface RibbonDisplayLayout {
	/** The title bar above the ribbon shows. */
	titleBar: boolean;
	/** The tab row (and the command row above it) shows. */
	tabs: boolean;
	/** The active tab's commands: in the layout, laid over the slide, or not shown. */
	commands: 'docked' | 'overlay' | 'hidden';
	/** The slim strip that brings the ribbon back shows (full-screen mode, nothing peeked). */
	reveal: boolean;
}

/**
 * What a binding draws for `mode`. `peek` is true while the user has called the commands up
 * temporarily: by choosing a tab in `tabs`, or by pressing the reveal strip in `fullScreen`.
 */
export function ribbonDisplayLayout(mode: RibbonDisplayMode, peek: boolean): RibbonDisplayLayout {
	if (mode === 'always') return { titleBar: true, tabs: true, commands: 'docked', reveal: false };
	if (mode === 'tabs')
		return { titleBar: true, tabs: true, commands: peek ? 'overlay' : 'hidden', reveal: false };
	return { titleBar: false, tabs: peek, commands: peek ? 'overlay' : 'hidden', reveal: !peek };
}
