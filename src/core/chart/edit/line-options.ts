import type { DrawingLine } from '../../drawingml/types';
import type { ChartSpace } from '../model';
import { editChartSpace } from './draft';
import type { ChartFormatTarget } from './target';
import { chartTargetDraftShapes } from './target-paint';

export type ChartLineEdit = {
	[Key in Exclude<keyof DrawingLine, 'sourceXml'>]?: DrawingLine[Key] | null;
};
const KEYS = new Set([
	'widthEmu',
	'fill',
	'dash',
	'cap',
	'compound',
	'alignment',
	'customDash',
	'join',
	'headEnd',
	'tailEnd',
]);
const ARROWS = ['none', 'triangle', 'stealth', 'diamond', 'oval', 'arrow'];

/** Null clears one property to Automatic without removing the rest of the authored outline. */
export function editChartLine(current: DrawingLine | undefined, edit: ChartLineEdit): DrawingLine {
	for (const key of Object.keys(edit))
		if (!KEYS.has(key)) throw new RangeError('Unknown chart outline setting');
	if (
		edit.widthEmu !== undefined &&
		edit.widthEmu !== null &&
		(!Number.isFinite(edit.widthEmu) || edit.widthEmu < 0)
	)
		throw new RangeError('Invalid chart outline width');
	if (edit.cap && !['flat', 'rnd', 'sq'].includes(edit.cap))
		throw new RangeError('Invalid chart outline cap');
	if (edit.compound && !['sng', 'dbl', 'thickThin', 'thinThick', 'tri'].includes(edit.compound))
		throw new RangeError('Invalid chart compound outline');
	if (edit.join && !['round', 'bevel', 'miter'].includes(edit.join.kind))
		throw new RangeError('Invalid chart outline join');
	for (const end of [edit.headEnd, edit.tailEnd])
		if (
			end &&
			((end.type && !ARROWS.includes(end.type)) ||
				(end.width && !['sm', 'med', 'lg'].includes(end.width)) ||
				(end.length && !['sm', 'med', 'lg'].includes(end.length)))
		)
			throw new RangeError('Invalid chart arrow end');
	const line = structuredClone(current ?? {});
	for (const key of Object.keys(edit) as Array<keyof ChartLineEdit>) {
		const value = edit[key];
		if (value === null) delete line[key];
		else if (value !== undefined) Object.assign(line, { [key]: structuredClone(value) });
	}
	if (edit.headEnd) line.headEnd = { ...current?.headEnd, ...edit.headEnd };
	if (edit.tailEnd) line.tailEnd = { ...current?.tailEnd, ...edit.tailEnd };
	return line;
}

export function setChartTargetLineOptions(
	space: ChartSpace,
	target: ChartFormatTarget,
	edit: ChartLineEdit,
): ChartSpace {
	return editChartSpace(space, (draft) => {
		for (const shape of chartTargetDraftShapes(draft, target))
			shape.line = editChartLine(shape.line, edit);
	});
}
