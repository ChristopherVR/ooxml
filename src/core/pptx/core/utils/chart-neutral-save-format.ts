import {
	applyChartAppearance,
	chartSeriesEntryAt,
	editChartSpace,
	parseChartFormatTargetKey,
} from '../../../chart';
import type { PptxChartData } from '../types';
import type { NeutralChartEdit } from './chart-neutral-save-plan';
import { same } from './chart-neutral-save-common';
import { mirrorFlatFields, mirrorTextColor } from '../builders/sdk/chart-element-format-mirror';
import { readChartElementFormat } from '../builders/sdk/chart-element-format-read';

/** Revert only preview fields that exactly match the format record's own derived changes. */
function claimMirrors(
	probe: Record<string, unknown>,
	before: Record<string, unknown>,
	mirror: Record<string, unknown>,
): void {
	for (const key of new Set([...Object.keys(before), ...Object.keys(mirror)])) {
		const old = before[key],
			value = mirror[key];
		if (same(old, value)) continue;
		if (same(probe[key], value)) {
			if (old === undefined) delete probe[key];
			else probe[key] = structuredClone(old);
		} else if (
			old &&
			value &&
			probe[key] &&
			typeof old === 'object' &&
			typeof value === 'object' &&
			typeof probe[key] === 'object'
		) {
			claimMirrors(
				probe[key] as Record<string, unknown>,
				old as Record<string, unknown>,
				value as Record<string, unknown>,
			);
		}
	}
}

/**
 * `edit` with the points of series `seriesIndex` kept as they are. A series fill set through the
 * shared operation clears the points' own fills, as Office's Format pane does, but the model
 * states point overrides on their own (`dataPoints`, written by the point family before this
 * one), so a series format keeps them, as the series colour does (`keepPointColors`).
 */
function keepingSeriesPoints(seriesIndex: number, edit: NeutralChartEdit): NeutralChartEdit {
	return (space) => {
		const points = chartSeriesEntryAt(space, seriesIndex).series.dataPoints;
		const next = edit(space);
		if (points.length === 0) return next;
		return editChartSpace(next, (draft) => {
			chartSeriesEntryAt(draft, seriesIndex).series.dataPoints = structuredClone(points);
		});
	};
}

/** The compatibility record is already shaped like the shared appearance contract. */
export function planNeutralChartFormats(
	before: PptxChartData,
	probe: PptxChartData,
	edits: NeutralChartEdit[],
): void {
	const old = before.elementFormats ?? {};
	const now = probe.elementFormats;
	if (!now) return;
	for (const [key, format] of Object.entries(now)) {
		if (same(old[key], format)) continue;
		const target = parseChartFormatTargetKey(key);
		if (!target || !format) continue;
		const source = old[key];
		const appearance = readChartElementFormat(before, target);
		const apply: NeutralChartEdit = (space) =>
			applyChartAppearance(space, target, appearance, format);
		edits.push(target.kind === 'series' ? keepingSeriesPoints(target.seriesIndex, apply) : apply);
		let mirror = mirrorFlatFields(before, target, format.fill, format.line);
		if (format.text) mirror = mirrorTextColor(mirror, target, format.text);
		claimMirrors(
			probe as unknown as Record<string, unknown>,
			before as unknown as Record<string, unknown>,
			mirror as unknown as Record<string, unknown>,
		);
		if (source === undefined) delete now[key];
		else now[key] = source;
	}
	if (!Object.keys(now).length && before.elementFormats === undefined) delete probe.elementFormats;
}
