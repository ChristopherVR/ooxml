import { elements } from '../xml/index';
import type { Cells, Report } from './sheet';
import { themeChild, themeChildren } from './theme-color';
import { themeFormat, type ThemeResources } from './theme-resolve';

/** DrawingML line widths are EMUs; one inch is 914400 EMUs. Cached values win. */
export function themeLineWeight(
	cells: Cells,
	resources: ThemeResources,
	report: Report,
): number | undefined {
	if (cells.get('LineWeight')?.value !== 'Themed') return undefined;
	const line = themeFormat(cells, 'Line', resources);
	const raw = line?.getAttribute('w');
	if (!line || !raw || !/^\d{1,8}$/.test(raw)) return undefined;
	const emu = Number(raw);
	if (emu > 20_116_800) return undefined;
	const compound = line.getAttribute('cmpd');
	if (compound && compound !== 'sng')
		report(
			'unsupported-theme-line-compound',
			'A compound theme line is approximated by a single stroke.',
		);
	return emu / 914_400;
}

/** Visio arrowhead codes with measured glyphs for the DrawingML end types (filled, stealth, open). */
const ARROW_CODES: Readonly<Record<string, number>> = {
	none: 0,
	triangle: 4,
	stealth: 5,
	arrow: 3,
};
const ARROW_SIZES: Readonly<Record<string, number>> = { sm: 1, med: 2, lg: 3 };

/**
 * A themed BeginArrow or EndArrow: the `a:headEnd` (begin) or `a:tailEnd` (end) of the selected
 * connector line style. Visio 16's own themes carry none, so the cells they write are literal;
 * this reads the extension for a theme that does. Diamond and oval ends have no measured glyph
 * and read as undefined.
 */
export function themeArrow(
	cells: Cells,
	resources: ThemeResources,
	side: 'Begin' | 'End',
): { code: number; size: number } | undefined {
	if (cells.get(`${side}Arrow`)?.value !== 'Themed') return undefined;
	const end = themeChildren(
		themeFormat(cells, 'Line', resources),
		side === 'Begin' ? 'headEnd' : 'tailEnd',
	)[0];
	if (!end) return { code: 0, size: 2 };
	const code = ARROW_CODES[end.getAttribute('type') ?? 'none'];
	return code === undefined
		? undefined
		: { code, size: ARROW_SIZES[end.getAttribute('w') ?? 'med'] ?? 2 };
}

/** Report only a selected nonempty effect style when the shape requests themed effects. */
export function reportThemeEffects(cells: Cells, resources: ThemeResources, report: Report): void {
	if (
		!['ShdwPattern', 'ShapeShdwShow', 'ReflectionTrans', 'GlowSize', 'SoftEdgesSize'].some(
			(name) => cells.get(name)?.value === 'Themed',
		)
	)
		return;
	const style = themeFormat(cells, 'Effects', resources);
	if (
		style &&
		elements(style).some(
			(node) =>
				node.localName === 'effectLst' &&
				elements(node).some(
					(effect) =>
						![
							'blur',
							'fillOverlay',
							'prstShdw',
							'outerShdw',
							'glow',
							'softEdge',
							'reflection',
						].includes(effect.localName),
				),
		)
	)
		report(
			'unsupported-theme-effects',
			'The selected theme effect style contains effects that are not rendered.',
		);
}

/**
 * ISO 29500 prstDash selects a named dashing scheme:
 * https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.presetdash
 * Only an explicit solid token is resolved here. Omission and non-solid spacing stay diagnosed.
 */
export function themeSolidLinePattern(
	cells: Cells,
	resources: ThemeResources,
): { linePattern: number } | undefined {
	if (cells.get('LinePattern')?.value !== 'Themed') return undefined;
	const line = themeFormat(cells, 'Line', resources);
	const dash = themeChildren(line, 'prstDash');
	if (dash.length !== 1 || themeChild(line, 'custDash') || dash[0]?.getAttribute('val') !== 'solid')
		return undefined;
	return { linePattern: 1 };
}
