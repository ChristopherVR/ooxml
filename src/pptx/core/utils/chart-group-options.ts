/**
 * Chart-group spacing and geometry carried on the chart-type container:
 * `c:gapWidth` and `c:overlap` (bar spacing), `c:firstSliceAng` (pie and
 * doughnut rotation) and `c:holeSize` (doughnut hole).
 *
 * One table drives both writers. A generated chart (`chart-xml-generator.ts`)
 * takes each value from the model, falling back to the constant it always
 * wrote (`gapWidth` 150, `holeSize` 50) or to nothing. A loaded chart is only
 * touched where the model and the authored element disagree, so an unedited
 * chart part round-trips byte-for-byte.
 *
 * Ranges are the `dml-chart.xsd` simple types; which container may hold which
 * element comes from `CHART_CONTAINER_CHILD_ORDER`.
 *
 * @module utils/chart-group-options
 */
import type { PptxChartData, XmlObject } from '../types';
import { chartContainerAllows } from './chart-container-content-model';
import { orderChartContainerChildren } from './chart-container-schema';

type GetLocalName = (key: string) => string;

/** One container option: its element, model field and schema range. */
export interface ChartGroupOptionSpec {
	element: 'gapWidth' | 'overlap' | 'firstSliceAng' | 'holeSize';
	field: 'barGapWidth' | 'barOverlap' | 'firstSliceAngle' | 'doughnutHoleSize';
	/** Schema simple type, for error messages. */
	schemaType: string;
	min: number;
	max: number;
	/** What a generated chart writes when the model leaves the field unset. */
	generatedDefault?: number;
}

/** The four options, with `dml-chart.xsd` ranges. */
export const CHART_GROUP_OPTIONS: readonly ChartGroupOptionSpec[] = [
	{
		element: 'gapWidth',
		field: 'barGapWidth',
		schemaType: 'ST_GapAmount',
		min: 0,
		max: 500,
		generatedDefault: 150,
	},
	{ element: 'overlap', field: 'barOverlap', schemaType: 'ST_Overlap', min: -100, max: 100 },
	{
		element: 'firstSliceAng',
		field: 'firstSliceAngle',
		schemaType: 'ST_FirstSliceAng',
		min: 0,
		max: 360,
	},
	{
		element: 'holeSize',
		field: 'doughnutHoleSize',
		schemaType: 'ST_HoleSize',
		min: 1,
		max: 90,
		generatedDefault: 50,
	},
];

/** `@val` for a model value: rounded and clamped into the schema range. */
export function chartGroupOptionVal(spec: ChartGroupOptionSpec, value: number): string {
	return String(Math.min(spec.max, Math.max(spec.min, Math.round(value))));
}

function localNameOf(key: string): string {
	return key.replace(/^.*:/u, '');
}

function findKey(obj: XmlObject, local: string, getLocalName: GetLocalName): string | undefined {
	return Object.keys(obj).find((key) => getLocalName(key) === local);
}

function modelValue(chartData: PptxChartData, spec: ChartGroupOptionSpec): number | undefined {
	if (spec.element === 'gapWidth' && chartData.chartType === 'ofPie') {
		return chartData.ofPieOptions?.gapWidth ?? chartData.barGapWidth;
	}
	const value = chartData[spec.field];
	return typeof value === 'number' && Number.isFinite(value) ? value : undefined;
}

/**
 * Write the options a freshly generated chart-type container allows, then put
 * the container back in schema order.
 *
 * @param containerLocal - The container's local name, e.g. `barChart`.
 */
export function applyGeneratedChartGroupOptions(
	container: XmlObject,
	containerLocal: string,
	chartData: PptxChartData,
): void {
	for (const spec of CHART_GROUP_OPTIONS) {
		if (!chartContainerAllows(containerLocal, spec.element)) {
			continue;
		}
		const value = modelValue(chartData, spec) ?? spec.generatedDefault;
		if (value !== undefined) {
			container[`c:${spec.element}`] = { '@_val': chartGroupOptionVal(spec, value) };
		}
	}
	orderChartContainerChildren(container, containerLocal, localNameOf);
}

function authoredValue(node: unknown): number | undefined {
	if (!node || typeof node !== 'object') {
		return undefined;
	}
	const raw = (node as XmlObject)['@_val'];
	const num =
		raw === undefined || raw === null || raw === '' ? NaN : Number.parseFloat(String(raw));
	return Number.isFinite(num) ? num : undefined;
}

/**
 * Reconcile a loaded chart's primary container with the model. An element is
 * rewritten only when the model holds a different value, and removed only
 * when the model cleared a value the parser had read from it; everything else
 * stays as authored. `ofPieChart` keeps `c:gapWidth` with the of-pie options
 * writer. The caller re-orders the container afterwards.
 *
 * @returns Whether the container changed.
 */
export function applyChartGroupOptionsToXml(
	container: XmlObject,
	containerLocal: string,
	chartData: PptxChartData,
	getLocalName: GetLocalName,
): boolean {
	let changed = false;
	for (const spec of CHART_GROUP_OPTIONS) {
		if (!chartContainerAllows(containerLocal, spec.element)) {
			continue;
		}
		if (spec.element === 'gapWidth' && containerLocal === 'ofPieChart') {
			continue;
		}
		const key = findKey(container, spec.element, getLocalName);
		const authored = key ? authoredValue(container[key]) : undefined;
		const value = chartData[spec.field];
		if (value === undefined) {
			if (key && authored !== undefined) {
				delete container[key];
				changed = true;
			}
			continue;
		}
		if (authored === value) {
			continue;
		}
		container[key ?? `c:${spec.element}`] = { '@_val': chartGroupOptionVal(spec, value) };
		changed = true;
	}
	return changed;
}
