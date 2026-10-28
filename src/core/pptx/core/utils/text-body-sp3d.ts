import { readText3DValues } from '../../../drawingml/read-3d-values';
import type { Text3DStyle, TextStyle, XmlObject } from '../types';

/**
 * Parse `a:bodyPr/a:sp3d` (§20.1.5.12) into {@link TextStyle.text3d}.
 *
 * Extracted from the body-properties parser so the absent-attribute handling
 * is testable in isolation. The guards used to read `attr !== null`, but
 * fast-xml-parser yields `undefined` (never `null`) for an attribute the
 * source never wrote - so `<a:sp3d><a:bevelT prst="circle"/></a:sp3d>`, which
 * is exactly how PowerPoint writes a default bevel, passed every guard and
 * stored `parseInt('undefined')` = `NaN` into `extrusionHeight`,
 * `bevelTopWidth`, `bevelTopHeight`, `bevelBottomWidth` and
 * `bevelBottomHeight`. A `NaN` reaches the renderers as a poisoned dimension
 * and serialises to `null` through JSON (the converter, the collaboration
 * codec), so it is corruption even where the save writer's truthiness check
 * happens to skip it.
 *
 * The shape-level twin (`applyShape3dStyle` in
 * `core/builders/shape-style-3d-helpers.ts`) already used the `!== undefined`
 * form; this brings the text-body path in line with it.
 *
 * @param bodyPr     - The `a:bodyPr` node to read.
 * @param style      - Text style to populate.
 * @param parseColor - Theme-aware colour resolver for `a:extrusionClr`.
 */
export function parseTextBodySp3d(
	bodyPr: XmlObject,
	style: TextStyle,
	parseColor: (colorNode: XmlObject | undefined) => string | undefined,
): void {
	// `a:sp3d` and `a:flatTx` are a mutually exclusive OOXML choice
	// (`EG_Text3D`): an explicit `<a:flatTx/>` is the authored signal that this
	// text body must NOT apply 3D (it overrides an inherited `sp3d`), so it is
	// modelled as its own boolean rather than left to silently fall out as "no
	// sp3d here". See `TextStyle.flatText`.
	if (bodyPr['a:flatTx'] !== undefined) {
		style.flatText = true;
		return;
	}

	const sp3d = bodyPr['a:sp3d'] as XmlObject | undefined;
	if (!sp3d) {
		return;
	}

	const values = readText3DValues(sp3d, {
		attribute: (node, name) => node?.[`@_${name}`],
		child: (node, name) => node?.[`a:${name}`] as XmlObject | undefined,
	});
	const { z: _z, contourWidth: _contourWidth, ...scalars } = values;
	const t3d = scalars as Text3DStyle;
	const extrusionColor = parseColor(sp3d['a:extrusionClr'] as XmlObject | undefined);
	if (extrusionColor) {
		t3d.extrusionColor = extrusionColor;
	}
	if (Object.keys(t3d).length > 0) {
		style.text3d = t3d;
	}
}
