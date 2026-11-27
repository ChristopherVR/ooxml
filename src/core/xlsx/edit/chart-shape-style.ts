// A Shape Styles gallery pick on a chart element (the chart Format tab): the theme styles resolve
// their style references against the workbook's format scheme, the presets carry their own fill
// and line, so the chart gets the explicit formatting Excel writes for the element.
import type { ChartShapeFormat } from '../../chart/edit/shape-format';
import type { DrawingOuterShadow } from '../../drawingml/effect-model';
import {
	resolveThemeEffectStyle,
	resolveThemeFillStyle,
	resolveThemeLineStyle,
} from '../../drawingml/format-scheme';
import {
	presetShapeStyle,
	themeShapeStyle,
	type ShapeStylePick,
} from '../../drawingml/shape-style-catalog';
import type { Workbook } from '../model';
import { officeFormatScheme } from '../office-format-scheme';

export function chartShapeStyleFormat(
	workbook: Workbook,
	pick: ShapeStylePick,
): ChartShapeFormat | undefined {
	if (pick.block === 'preset') {
		const preset = presetShapeStyle(pick.row, pick.column);
		return preset && { fill: preset.fill, line: preset.line, shadow: null };
	}
	const refs = themeShapeStyle(pick.row, pick.column);
	if (!refs) return undefined;
	const scheme = workbook.theme?.formatScheme ?? officeFormatScheme();
	const fill =
		refs.fill.idx > 0 ? resolveThemeFillStyle(scheme, refs.fill.idx, refs.fill.color) : undefined;
	const line =
		refs.line.idx > 0 ? resolveThemeLineStyle(scheme, refs.line.idx, refs.line.color) : undefined;
	const shadow =
		refs.effect.idx > 0
			? resolveThemeEffectStyle(scheme, refs.effect.idx, refs.effect.color)?.effects.find(
					(effect): effect is DrawingOuterShadow => effect.kind === 'outerShadow',
				)
			: undefined;
	return {
		fill: fill ?? { kind: 'none' },
		line: line ?? { fill: { kind: 'none' } },
		shadow: shadow ?? null,
	};
}
