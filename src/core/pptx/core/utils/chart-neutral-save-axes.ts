import { setChartAxis, type ChartAxisEdit } from '../../../chart/index';
import type { PptxChartAxisFormatting, PptxChartData } from '../types';
import type { NeutralChartEdit } from './chart-neutral-save-plan';
import { same } from './chart-neutral-save-common';

/** The axis fields the shared `setChartAxis` writes as the PowerPoint axis writers do. */
function axisEdit(
	was: PptxChartAxisFormatting,
	now: PptxChartAxisFormatting,
): { edit: ChartAxisEdit; claimed: (keyof PptxChartAxisFormatting)[] } {
	const edit: ChartAxisEdit = {};
	const claimed: (keyof PptxChartAxisFormatting)[] = [];
	for (const [source, target] of [
		['axPos', 'position'],
		['crossBetween', 'crossBetween'],
	] as const)
		if (was[source] !== now[source]) {
			Object.assign(edit, { [target]: now[source] ?? null });
			claimed.push(source);
		}
	if (was.crossAxisId !== now.crossAxisId && now.crossAxisId !== undefined) {
		edit.crossAxisId = now.crossAxisId;
		claimed.push('crossAxisId');
	}
	if (
		(was.crosses !== now.crosses || was.crossesAt !== now.crossesAt) &&
		(now.crossesAt !== undefined || now.crosses !== undefined)
	) {
		edit.crosses = now.crossesAt ?? now.crosses;
		claimed.push('crosses', 'crossesAt');
	}
	for (const key of ['min', 'max', 'majorUnit', 'minorUnit'] as const) {
		if (was[key] === now[key]) continue;
		edit[key] = now[key] ?? null;
		claimed.push(key);
	}
	if (Boolean(was.deleted) !== Boolean(now.deleted)) {
		edit.visible = !now.deleted;
		claimed.push('deleted');
	}
	if (!same(was.numFmt, now.numFmt) && now.numFmt) {
		edit.numberFormat = now.numFmt.formatCode;
		edit.numberFormatSourceLinked = now.numFmt.sourceLinked ?? false;
		claimed.push('numFmt');
	}
	if (was.titleText !== now.titleText) {
		edit.title = now.titleText ?? null;
		claimed.push('titleText');
	}
	for (const key of ['majorGridlines', 'minorGridlines', 'auto', 'noMultiLevelLabels'] as const) {
		if (was[key] === now[key] || now[key] === undefined) continue;
		edit[key] = now[key];
		claimed.push(key);
	}
	for (const key of ['majorTickMark', 'minorTickMark'] as const) {
		if (was[key] === now[key] || now[key] === undefined) continue;
		edit[key] = now[key];
		claimed.push(key);
	}
	for (const key of ['majorTimeUnit', 'minorTimeUnit', 'baseTimeUnit'] as const) {
		if (was[key] === now[key]) continue;
		edit[key] = now[key] ?? null;
		claimed.push(key);
	}
	for (const key of ['labelOffset', 'tickLabelSkip', 'tickMarkSkip'] as const) {
		if (was[key] === now[key]) continue;
		edit[key] = now[key] ?? null;
		claimed.push(key);
	}
	if (was.labelAlignment !== now.labelAlignment) {
		edit.labelAlign = now.labelAlignment ?? null;
		claimed.push('labelAlignment');
	}
	if (was.logScale !== now.logScale || was.logBase !== now.logBase) {
		edit.logBase =
			now.logScale === false || (now.logScale !== true && now.logBase === undefined)
				? null
				: (now.logBase ?? 10);
		claimed.push('logScale', 'logBase');
	}
	if (was.displayUnits !== now.displayUnits || was.displayUnitsValue !== now.displayUnitsValue) {
		edit.displayUnits =
			now.displayUnits === 'custom' ? now.displayUnitsValue : (now.displayUnits ?? null);
		if (edit.displayUnits !== undefined) claimed.push('displayUnits', 'displayUnitsValue');
	}
	if (was.orientation !== now.orientation && now.orientation !== undefined) {
		edit.reverse = now.orientation === 'maxMin';
		claimed.push('orientation');
	}
	if (was.tickLblPos !== now.tickLblPos && now.tickLblPos !== undefined) {
		edit.tickLabelPosition = now.tickLblPos;
		claimed.push('tickLblPos');
	}
	return { edit, claimed };
}

/** Value and category axis scaling, visibility, number format, direction and labels. */
export function planAxes(
	before: PptxChartData,
	probe: PptxChartData,
	edits: NeutralChartEdit[],
): void {
	const was = before.axes ?? [];
	const now = probe.axes ?? [];
	if (was.length !== now.length) return;
	now.forEach((axis, i) => {
		const previous = was[i]!;
		const id = previous.axisId;
		if (id === undefined || axis.axisId !== id || axis.axisType !== previous.axisType) return;
		const { edit, claimed } = axisEdit(previous, axis);
		if (claimed.length === 0) return;
		edits.push((space) => setChartAxis(space, { id }, edit));
		const reverted: Record<string, unknown> = { ...axis };
		for (const key of claimed) {
			if (previous[key] === undefined) delete reverted[key];
			else reverted[key] = previous[key];
		}
		now[i] = reverted as unknown as PptxChartAxisFormatting;
	});
}
