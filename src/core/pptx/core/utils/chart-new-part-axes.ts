/**
 * The axes of a new classic chart part written through the shared chart core
 * (`chart-new-part.ts`). The part has the axes the legacy generator always
 * wrote, with its ids (`GENERATED_CHART_AXIS_IDS`), and a model axis is
 * matched to one the way the generator matched it: by axis id, else by type
 * and order (the first `catAx`, `valAx`, `dateAx` or `serAx`; for scatter and
 * bubble charts the first and second `valAx`). SDK axes carry no id
 * (`setChartAxis` adds `{ axisType }` entries), so the ids given here are what
 * lets the save planner's axis family claim their edits.
 *
 * A model value axis no slot takes is a secondary axis when a series is plotted
 * on it and it crosses another unmatched axis; the series is moved onto it
 * with `setChartSeriesSecondaryAxis`, keeping the model's ids. A model axis the
 * part has no place for (a pie chart's, or a pair no series is plotted on) is
 * not written and is reported.
 *
 * @module utils/chart-new-part-axes
 */
import type { PptxChartAxisFormatting, PptxChartData, PptxChartElementFormats } from '../types';
import { GENERATED_CHART_AXIS_IDS } from './chart-xml-container-map';

type AxisType = PptxChartAxisFormatting['axisType'];
type AxisPosition = NonNullable<PptxChartAxisFormatting['axPos']>;

interface AxisSlot {
	type: AxisType;
	id: number;
	cross: number;
	position: AxisPosition;
	/** Which axis of `type` the slot takes when no model axis has its id. */
	typeIndex: number;
}

/** A series the part plots on a secondary axis pair, with the pair's ids. */
export interface NewChartSecondaryAxis {
	seriesIndex: number;
	/** The secondary category and value axis ids. */
	axisIds: [number, number];
}

export interface NewChartAxes {
	/** The model's axes in the part's order, ids, crossings and positions filled in. */
	after: PptxChartAxisFormatting[];
	/** What the part holds for each of `after` before the planner's edits. */
	baseline: PptxChartAxisFormatting[];
	/** Series to move onto a secondary axis pair, in series order. */
	secondary: NewChartSecondaryAxis[];
	/** Whether the category axis is a date axis. */
	dateAxis: boolean;
	/** `elementFormats` with the axis-indexed keys moved to the part's axis order. */
	elementFormats?: PptxChartElementFormats;
	/** Model fields the part has no place for, as paths. */
	unwritten: string[];
}

const PIE_FAMILY = new Set(['pie', 'pie3D', 'doughnut', 'ofPie']);
const XY = new Set(['scatter', 'bubble']);
const AXIS_FORMAT_KEY = /^(axis|majorGridlines|minorGridlines):(\d+)$/u;

function slotsOf(chart: PptxChartData, model: readonly PptxChartAxisFormatting[]): AxisSlot[] {
	if (PIE_FAMILY.has(chart.chartType)) return [];
	const { cat, val, ser } = GENERATED_CHART_AXIS_IDS;
	if (XY.has(chart.chartType))
		return [
			{ type: 'valAx', id: cat, cross: val, position: 'b', typeIndex: 0 },
			{ type: 'valAx', id: val, cross: cat, position: 'l', typeIndex: 1 },
		];
	const horizontal = chart.barDirection === 'bar' && ['bar', 'bar3D'].includes(chart.chartType);
	const first: AxisSlot = {
		type: 'catAx',
		id: cat,
		cross: val,
		position: horizontal ? 'l' : 'b',
		typeIndex: 0,
	};
	const date = { ...first, type: 'dateAx' as const };
	const dated = chart.dateCategories !== undefined && matchOf(model, date) !== undefined;
	return [
		dated ? date : first,
		{ type: 'valAx', id: val, cross: cat, position: horizontal ? 'b' : 'l', typeIndex: 0 },
		...(chart.chartType === 'line3D' || chart.chartType === 'surface'
			? [{ type: 'serAx', id: ser, cross: val, position: 'b', typeIndex: 0 } as const]
			: []),
	];
}

const matchOf = (axes: readonly PptxChartAxisFormatting[], slot: AxisSlot) =>
	axes.find((axis) => axis.axisId === slot.id) ??
	axes.filter((axis) => axis.axisType === slot.type)[slot.typeIndex];

/** The secondary pairs series are plotted on, among the axes no slot took. */
function secondaryPairs(
	chart: PptxChartData,
	extras: readonly PptxChartAxisFormatting[],
): {
	pairs: [PptxChartAxisFormatting, PptxChartAxisFormatting][];
	series: NewChartSecondaryAxis[];
} {
	const pairs: [PptxChartAxisFormatting, PptxChartAxisFormatting][] = [];
	const series: NewChartSecondaryAxis[] = [];
	chart.series.forEach((item, seriesIndex) => {
		if (item.axisId === undefined) return;
		const value = extras.find((axis) => axis.axisType === 'valAx' && axis.axisId === item.axisId);
		const cross =
			value?.crossAxisId === undefined
				? undefined
				: extras.find((axis) => axis !== value && axis.axisId === value.crossAxisId);
		if (!value || !cross || cross.axisId === undefined) return;
		if (!pairs.some(([, known]) => known === value)) pairs.push([cross, value]);
		series.push({ seriesIndex, axisIds: [cross.axisId, value.axisId!] });
	});
	return { pairs, series };
}

/** The stub of a secondary pair as `setChartSeriesSecondaryAxis` writes it. */
function secondaryStubs(
	cross: PptxChartAxisFormatting,
	value: PptxChartAxisFormatting,
): PptxChartAxisFormatting[] {
	return [
		{
			axisType: cross.axisType,
			axisId: cross.axisId!,
			crossAxisId: value.axisId!,
			axPos: 't',
			deleted: true,
		},
		{
			axisType: 'valAx',
			axisId: value.axisId!,
			crossAxisId: cross.axisId!,
			axPos: 'r',
			crosses: 'max',
		},
	];
}

function movedFormats(
	formats: PptxChartElementFormats | undefined,
	moved: ReadonlyMap<number, number>,
	unwritten: string[],
): PptxChartElementFormats | undefined {
	if (!formats) return undefined;
	const out: PptxChartElementFormats = {};
	for (const [key, value] of Object.entries(formats)) {
		const match = AXIS_FORMAT_KEY.exec(key);
		if (!match) {
			out[key] = value;
			continue;
		}
		const index = moved.get(Number(match[2]));
		if (index === undefined) unwritten.push(`elementFormats.${key}`);
		else out[`${match[1]}:${index}`] = value;
	}
	return out;
}

/**
 * `chart`'s axes with the ids the new part gives them, before its first save: each model axis
 * the part matches by type and order takes the generated id, crossing and position it lacks, and
 * a generated axis no model axis stands for is appended. The order of the model's axes is kept, so
 * axis-indexed element formats still name the same axes. A chart whose axes already have ids, or
 * whose type has no axes, is returned as it is.
 */
export function withGeneratedAxisIds(chart: PptxChartData): PptxChartData {
	const model = chart.axes ?? [];
	if (model.some((axis) => axis.axisId !== undefined)) return chart;
	const slots = slotsOf(chart, model);
	if (slots.length === 0) return chart;
	const axes = model.map((axis) => structuredClone(axis));
	const taken = new Set<PptxChartAxisFormatting>();
	for (const slot of slots) {
		const own = axes.filter((axis) => axis.axisType === slot.type)[slot.typeIndex];
		if (own && !taken.has(own)) {
			taken.add(own);
			own.axisId = slot.id;
			own.crossAxisId ??= slot.cross;
			own.axPos ??= slot.position;
			continue;
		}
		axes.push({
			axisType: slot.type,
			axisId: slot.id,
			crossAxisId: slot.cross,
			axPos: slot.position,
		});
	}
	return { ...chart, axes };
}

/** The axes of the new part for `chart`. */
export function newChartAxes(chart: PptxChartData): NewChartAxes {
	const model = chart.axes ?? [];
	const slots = slotsOf(chart, model);
	const matched = slots.map((slot) => matchOf(model, slot));
	const taken = new Set(matched.filter((axis) => axis !== undefined));
	const extras = model.filter((axis) => !taken.has(axis));
	const { pairs, series } = secondaryPairs(chart, extras);
	const after: PptxChartAxisFormatting[] = [];
	const baseline: PptxChartAxisFormatting[] = [];
	const moved = new Map<number, number>();
	slots.forEach((slot, i) => {
		const stub: PptxChartAxisFormatting = {
			axisType: slot.type,
			axisId: slot.id,
			crossAxisId: slot.cross,
			axPos: slot.position,
		};
		const own = matched[i];
		if (own) moved.set(model.indexOf(own), after.length);
		baseline.push(stub);
		after.push(
			own
				? {
						...structuredClone(own),
						axisId: own.axisId ?? slot.id,
						crossAxisId: own.crossAxisId ?? slot.cross,
						axPos: own.axPos ?? slot.position,
					}
				: stub,
		);
	});
	for (const [cross, value] of pairs) {
		const stubs = secondaryStubs(cross, value);
		[cross, value].forEach((own, i) => {
			moved.set(model.indexOf(own), after.length);
			baseline.push(stubs[i]!);
			after.push(structuredClone(own));
		});
	}
	const unwritten = model.flatMap((axis, index) =>
		moved.has(index) ? [] : [`axes.${index}.${axis.axisType}`],
	);
	const elementFormats = movedFormats(chart.elementFormats, moved, unwritten);
	return {
		after,
		baseline,
		secondary: series,
		dateAxis: slots[0]?.type === 'dateAx',
		...(elementFormats ? { elementFormats } : {}),
		unwritten,
	};
}
