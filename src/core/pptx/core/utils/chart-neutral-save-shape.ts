import {
	chartDrawingColorFromHex,
	editChartShapeProperties,
	editChartSpace,
	chartTargetDraftShapes,
	type ChartShapeEdit,
	type ChartFormatTarget,
} from '../../../chart';
import type { PptxChartData, PptxChartShapeProps } from '../types';
import type { NeutralChartEdit } from './chart-neutral-save-plan';
import { same, revertStyle } from './chart-neutral-save-common';

function fill(value: string | undefined) {
	if (value === 'none') return { kind: 'none' as const };
	if (!value) return null;
	const color = chartDrawingColorFromHex(value);
	if (!color) throw new RangeError('Invalid chart shape colour');
	return { kind: 'solid' as const, color };
}

export function neutralShapeEdit(
	before: PptxChartShapeProps = {},
	after: PptxChartShapeProps = {},
): ChartShapeEdit {
	const edit: ChartShapeEdit = {};
	if (before.fillColor !== after.fillColor) edit.fill = fill(after.fillColor);
	if (before.strokeColor !== after.strokeColor || before.lineNoFill !== after.lineNoFill)
		edit.line = { fill: after.lineNoFill ? { kind: 'none' } : fill(after.strokeColor) };
	if (before.strokeWidth !== after.strokeWidth) {
		edit.line ??= {};
		edit.line.widthEmu =
			after.strokeWidth === undefined ? null : Math.round(after.strokeWidth * 12700);
	}
	if (before.strokeDashStyle !== after.strokeDashStyle) {
		edit.line ??= {};
		edit.line.dash = after.strokeDashStyle ?? null;
	}
	return edit;
}

function shapeEdit(
	edits: NeutralChartEdit[],
	target: ChartFormatTarget,
	before: PptxChartShapeProps | null | undefined,
	after: PptxChartShapeProps | null | undefined,
) {
	const edit = neutralShapeEdit(before ?? {}, after ?? {});
	edits.push((space) =>
		editChartSpace(space, (draft) => {
			for (const shape of chartTargetDraftShapes(draft, target)) {
				if (after === null)
					for (const key of Object.keys(shape)) delete (shape as Record<string, unknown>)[key];
				else Object.assign(shape, editChartShapeProperties(shape, edit));
			}
		}),
	);
}

export function planNeutralChartShapes(
	before: PptxChartData,
	probe: PptxChartData,
	edits: NeutralChartEdit[],
): void {
	if (!same(before.style?.titleSpPr, probe.style?.titleSpPr)) {
		shapeEdit(edits, { kind: 'title' }, before.style?.titleSpPr, probe.style?.titleSpPr);
		revertStyle(probe, before, ['titleSpPr']);
	}
	for (const [index, axis] of (probe.axes ?? []).entries()) {
		const old = before.axes?.[index];
		if (!old || old.axisId !== axis.axisId) continue;
		for (const [key, kind] of [
			['spPr', 'axis'],
			['majorGridlinesSpPr', 'majorGridlines'],
			['minorGridlinesSpPr', 'minorGridlines'],
		] as const) {
			if (same(old[key], axis[key])) continue;
			shapeEdit(edits, { kind, axisIndex: index }, old[key], axis[key]);
			if (old[key] === undefined) delete axis[key];
			else axis[key] = old[key];
		}
	}
}
