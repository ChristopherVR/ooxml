import {
	visioFallbackQuickStyle,
	type VisioPage,
	type VisioQuickStyleColor,
} from 'ooxml-core/visio';
import type { OfficeUiColorGrid } from '../controls';
import type { VisioFormattingAction } from './ribbon-action';

/** What a colour menu paints: a shape's fill, its line, or its text. */
export type ColorTarget = 'fill' | 'line' | 'font';

const GRIDS: Readonly<Record<ColorTarget, { label: string; none?: string }>> = {
	fill: { label: 'Fill colors', none: 'No Fill' },
	line: { label: 'Line colors', none: 'No Line' },
	font: { label: 'Font colors' },
};
export const MORE_COLORS = 'More Colors...';

/**
 * Office's colour picker for a ribbon menu or the Format Shape pane: Theme Colors, Standard
 * Colors, recent colours, No Fill or No Line, and More Colors. `data-color-grid` names the target.
 */
export function colorGrid(doc: Document, target: ColorTarget): OfficeUiColorGrid {
	const grid = doc.createElement('office-ui-color-grid') as OfficeUiColorGrid;
	grid.dataset.colorGrid = target;
	// The menu's arrow keys and its opening focus go through the grid.
	grid.dataset.menuItem = '';
	grid.setAttribute('label', GRIDS[target].label);
	if (GRIDS[target].none) grid.setAttribute('none-label', GRIDS[target].none!);
	grid.setAttribute('more-label', MORE_COLORS);
	return grid;
}

/** The formatting action a picked colour (`#rrggbb` or `'none'`) stands for. */
export function colorAction(target: ColorTarget, color: string): VisioFormattingAction {
	if (target === 'font') return { type: 'font-color', value: color };
	if (target === 'fill') return { type: 'shape-format', patch: { fillColor: color } };
	return color === 'none'
		? { type: 'shape-format', patch: { linePattern: 0 } }
		: { type: 'shape-format', patch: { lineColor: color } };
}

/**
 * The ten Theme Colors columns of a page. A Visio page theme carries its six accents; the
 * background and text columns are Office's (left undefined, so the shared palette fills them
 * in). A drawing without a theme gets the Office accents Quick Styles also falls back to.
 */
export function pageThemeColors(page: VisioPage | undefined): (string | undefined)[] {
	return [
		undefined,
		undefined,
		undefined,
		undefined,
		...[1, 2, 3, 4, 5, 6].map(
			(accent) =>
				page?.theme?.accents[accent - 1] ??
				visioFallbackQuickStyle({ color: (accent + 1) as VisioQuickStyleColor, matrix: 4 }).fill,
		),
	];
}

/** The value all `values` share, or undefined when they differ or there are none. */
export function commonColor(values: readonly (string | undefined)[]): string | undefined {
	const first = values[0]?.toLowerCase();
	return first !== undefined && values.every((value) => value?.toLowerCase() === first)
		? first
		: undefined;
}
