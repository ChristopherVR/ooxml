import type { ShapeStyle, XmlObject } from '../../types';

/**
 * The `a:duotone` colour pair of an effect DAG. The other DAG colour operations are read by the
 * shared interpretation (`drawingml/effect-dag.ts`); a duotone's colours resolve here through the
 * runtime colour codec, against the slide's theme.
 */
export interface DagSpecificContext {
	parseColor: (colorNode: XmlObject | undefined, placeholderColor?: string) => string | undefined;
	ensureArray: (value: unknown) => XmlObject[];
}

export function extractDagDuotone(
	dag: XmlObject,
	style: Partial<ShapeStyle>,
	context: DagSpecificContext,
): void {
	const duotone = dag['a:duotone'] as XmlObject | undefined;
	if (!duotone) {
		return;
	}

	const colorNodes: XmlObject[] = [
		...context.ensureArray(duotone['a:srgbClr']),
		...context.ensureArray(duotone['a:schemeClr']),
		...context.ensureArray(duotone['a:prstClr']),
	];
	if (colorNodes.length >= 2) {
		style.dagDuotone = {
			color1: context.parseColor(colorNodes[0]) || '#000000',
			color2: context.parseColor(colorNodes[1]) || '#ffffff',
		};
	}
}
