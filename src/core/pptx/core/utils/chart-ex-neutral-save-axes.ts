import { editChartExAxis, type ChartExAxisEdit } from '../../../chart';
import type { PptxChartData, PptxChartAxisFormatting } from '../types';
import type { ChartExEdit } from './chart-ex-neutral-save';
import { same } from './chart-neutral-save-common';
const divisors: Record<string, string> = {
	'100': 'hundreds',
	'1000': 'thousands',
	'10000': 'tenThousands',
	'100000': 'hundredThousands',
	'1000000': 'millions',
	'10000000': 'tenMillions',
	'100000000': 'hundredMillions',
	'1000000000': 'billions',
	'1000000000000': 'trillions',
};
export function planChartExAxes(
	before: PptxChartData,
	probe: PptxChartData,
	edits: ChartExEdit[],
): void {
	for (const [index, axis] of (probe.axes ?? []).entries()) {
		const old = before.axes?.[index];
		if (!old || old.axisId !== axis.axisId) continue;
		const edit: ChartExAxisEdit = {},
			claimed: (keyof PptxChartAxisFormatting)[] = [];
		for (const key of ['min', 'max', 'majorUnit', 'minorUnit'] as const)
			if (old[key] !== axis[key]) {
				edit[key] = axis[key] ?? null;
				claimed.push(key);
			}
		if (old.deleted !== axis.deleted) {
			edit.visible = !axis.deleted;
			claimed.push('deleted');
		}
		if (old.titleText !== axis.titleText) {
			edit.title = axis.titleText ?? null;
			claimed.push('titleText');
		}
		if (!same(old.numFmt, axis.numFmt)) {
			edit.numberFormat = axis.numFmt ?? null;
			claimed.push('numFmt');
		}
		for (const key of ['majorGridlines', 'minorGridlines'] as const)
			if (old[key] !== axis[key]) {
				edit[key] = axis[key] ?? false;
				claimed.push(key);
			}
		for (const key of ['majorTickMark', 'minorTickMark'] as const)
			if (old[key] !== axis[key]) {
				edit[key] = axis[key] ?? null;
				claimed.push(key);
			}
		if (old.tickLblPos !== axis.tickLblPos) {
			if (axis.tickLblPos && !['none', 'nextTo'].includes(axis.tickLblPos))
				edits.push(() => {
					throw new RangeError('ChartEx axes only support showing or hiding tick labels.');
				});
			edit.tickLabels = axis.tickLblPos !== 'none';
			claimed.push('tickLblPos');
		}
		if (
			old.displayUnits !== axis.displayUnits ||
			old.displayUnitsValue !== axis.displayUnitsValue
		) {
			edit.units =
				axis.displayUnits === 'custom'
					? (divisors[String(axis.displayUnitsValue)] ?? 'unsupported-custom-unit')
					: (axis.displayUnits ?? null);
			claimed.push('displayUnits', 'displayUnitsValue');
		}
		for (const key of [
			'logScale',
			'logBase',
			'orientation',
			'crosses',
			'crossesAt',
			'crossBetween',
		] as const)
			if (!same(old[key], axis[key])) {
				edits.push(() => {
					throw new RangeError(`ChartEx does not support ${key} axis edits.`);
				});
				claimed.push(key);
			}
		if (claimed.length) edits.push((space) => editChartExAxis(space, index, edit));
		for (const key of claimed)
			if (old[key] === undefined) delete axis[key];
			else Object.assign(axis, { [key]: old[key] });
	}
}
