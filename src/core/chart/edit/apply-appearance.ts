import type { ChartSpace } from '../model';
import type {
	ChartEffectsAppearance,
	ChartFillAppearance,
	ChartLineAppearance,
	ChartTextAppearance,
	ChartThreeDAppearance,
} from './appearance-types';
import type { ChartFormatTarget } from './target';
import { chartTargetShape, chartTargetDraftShapes, setChartTargetFill } from './target-paint';
import { setChartTargetLineOptions } from './line-options';
import {
	chartEffectEditsFromAppearance,
	chartTextEditsFromAppearance,
	chartThreeDFromAppearance,
} from './format-values';
import { fillAppearancePatch, lineAppearancePatch, sameAppearance } from './appearance-fill';
import { setChartTargetEffect, type ChartEffectEdit } from './effects';
import { setChartTargetText } from './text';
import { setChartTargetThreeD } from './three-d';
import { editChartSpace } from './draft';

export interface ChartAppearance {
	reset?: boolean;
	fill?: ChartFillAppearance;
	line?: ChartLineAppearance;
	effects?: ChartEffectsAppearance;
	threeD?: ChartThreeDAppearance;
	text?: ChartTextAppearance;
}

function effects(
	before: ChartEffectsAppearance | undefined,
	after: ChartEffectsAppearance,
): ChartEffectEdit[] {
	const changed: ChartEffectsAppearance = {};
	for (const key of ['shadow', 'glow', 'softEdge'] as const)
		if (!sameAppearance(before?.[key], after[key]))
			Object.assign(changed, { [key]: after[key] ?? null });
	const edits = chartEffectEditsFromAppearance(changed);
	for (const [field, kind] of [
		['shadow', 'outerShadow'],
		['glow', 'glow'],
	] as const) {
		const old = before?.[field],
			next = after[field];
		if (!old || !next || old.color !== next.color) continue;
		for (const edit of edits) if (edit.kind === kind && 'color' in edit) delete edit.color;
		if (old.transparency !== next.transparency)
			edits.push({ kind, transparency: next.transparency ?? 0 });
	}
	return edits;
}

/** Apply differences from a host's resolved appearance without pinning unchanged theme colours. */
export function applyChartAppearance(
	space: ChartSpace,
	target: ChartFormatTarget,
	before: ChartAppearance | undefined,
	after: ChartAppearance,
): ChartSpace {
	let next = space;
	if (after.reset && !before?.reset) {
		next = editChartSpace(next, (draft) => {
			for (const shape of chartTargetDraftShapes(draft, target))
				for (const key of Object.keys(shape)) delete (shape as Record<string, unknown>)[key];
		});
		before = undefined;
	}
	if (after.fill && !sameAppearance(before?.fill, after.fill))
		next = setChartTargetFill(
			next,
			target,
			fillAppearancePatch(chartTargetShape(next, target)?.fill, before?.fill, after.fill),
		);
	if (after.line && !sameAppearance(before?.line, after.line))
		next = setChartTargetLineOptions(
			next,
			target,
			lineAppearancePatch(chartTargetShape(next, target)?.line, before?.line, after.line),
		);
	if (after.effects && !sameAppearance(before?.effects, after.effects))
		for (const edit of effects(before?.effects, after.effects))
			next = setChartTargetEffect(next, target, edit);
	if (after.threeD && !sameAppearance(before?.threeD, after.threeD)) {
		const patch: ChartThreeDAppearance = {};
		for (const key of Object.keys(after.threeD) as (keyof ChartThreeDAppearance)[])
			if (!sameAppearance(before?.threeD?.[key], after.threeD[key]))
				Object.assign(patch, { [key]: after.threeD[key] });
		const edit = chartThreeDFromAppearance(patch);
		for (const [key, color] of [
			['depth', 'depthColor'],
			['contour', 'contourColor'],
		] as const)
			if (before?.threeD?.[key]?.color === after.threeD[key]?.color && after.threeD[key])
				delete edit[color];
		next = setChartTargetThreeD(next, target, edit);
	}
	if (after.text && !sameAppearance(before?.text, after.text)) {
		const patch: ChartTextAppearance = {};
		for (const key of Object.keys(after.text) as (keyof ChartTextAppearance)[])
			if (!sameAppearance(before?.text?.[key], after.text[key]))
				Object.assign(patch, { [key]: after.text[key] });
		const { effects: textEffects, ...textPatch } = patch;
		const edits = chartTextEditsFromAppearance(textPatch);
		if (textEffects)
			edits.push(...effects(before?.text?.effects, textEffects).map((effect) => ({ effect })));
		for (const edit of edits) {
			if (edit.fill !== undefined && after.text.fill) {
				delete edit.fill;
				edit.fillAppearance = {
					...(before?.text?.fill ? { before: before.text.fill } : {}),
					after: after.text.fill,
				};
			}
			if (edit.line !== undefined && after.text.outline) {
				delete edit.line;
				edit.lineAppearance = {
					...(before?.text?.outline ? { before: before.text.outline } : {}),
					after: after.text.outline,
				};
			}
			next = setChartTargetText(next, target, edit);
		}
	}
	return next;
}
