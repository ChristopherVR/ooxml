import { elements } from '../xml/index.js';
import type { Cells, Report } from './sheet.js';
import { themeChild, themeChildren } from './theme-color.js';
import { themeFormat, type ThemeResources } from './theme-resolve.js';

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
					(effect) => !['blur', 'fillOverlay', 'prstShdw'].includes(effect.localName),
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
