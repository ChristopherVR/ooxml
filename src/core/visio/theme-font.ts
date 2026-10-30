/**
 * Theme fonts (MS-VSDX FontSchemeIndex, 2.4.4.58). A character Font cell of `Themed` takes the
 * body (minor) Latin typeface of the theme part whose font scheme ID the shape selects, or the
 * page when the shape inherits (65534). Checked against Microsoft Visio, which draws themed text
 * in the theme's body font for shapes and connectors alike.
 */
import type { Cells } from './sheet';
import { integer } from './theme-color';
import type { ThemeResources } from './theme-resolve';

const index = (cells: Cells | undefined) => integer(cells?.get('FontSchemeIndex')?.value);

/** The font scheme ID in force for a shape, or undefined when none is selected. */
export function selectedFontScheme(cells: Cells, resources: ThemeResources): number | undefined {
	const local = index(cells);
	const selected = local !== undefined && local !== 65534 ? local : index(resources.pageCells);
	return selected === undefined || selected === 0 || selected >= 65534 ? undefined : selected;
}

/** The body font of the one theme part carrying the selected font scheme ID. */
export function themeFontFamily(cells: Cells, resources: ThemeResources): string | undefined {
	const id = selectedFontScheme(cells, resources);
	if (id === undefined) return undefined;
	const matches = (resources.themes ?? []).filter((theme) => theme.fontId === id);
	return matches.length === 1 ? matches[0]!.minorFont : undefined;
}
