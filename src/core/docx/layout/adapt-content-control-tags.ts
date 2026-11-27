import type { ContentControl, TextRun } from '../model';
import type { LayoutRun } from './input';
import type { BlockContentControlTags } from './block-content-control-tags';
import { contentControlTagMetrics } from './content-control-tag-metrics';

function tag(control: ContentControl, edge: 'start' | 'end'): LayoutRun {
	const structural = control.level !== 'inline';
	return {
		text: edge === 'start' || structural ? (control.alias ?? '').replace(/\s+/g, ' ').trim() : '',
		contentControlTag: edge,
		...(structural && { blockContentControlTag: true }),
		synthetic: true,
		sourceLength: 0,
		fontSizePt: contentControlTagMetrics.fontSizePt,
		fontFamily: contentControlTagMetrics.fontFamily,
	};
}

/** Adds the enclosing controls' tags outside any inline tags already in the paragraph. */
export function adaptBlockContentControlTags(
	runs: LayoutRun[],
	edges: BlockContentControlTags,
): LayoutRun[] {
	return [
		...edges.start.map((control) => tag(control, 'start')),
		...runs,
		...edges.end.map((control) => tag(control, 'end')),
	];
}

/** A span can include several source runs and each can produce several font runs. */
export function adaptContentControlTags(
	runs: readonly TextRun[],
	adapt: (run: TextRun) => LayoutRun[],
): LayoutRun[] {
	const result: LayoutRun[] = [];
	let active: ContentControl | undefined;
	let hasVisibleContent = false;
	for (const run of runs) {
		const control = run.contentControl?.appearance === 'tags' ? run.contentControl : undefined;
		if (active && control?.id !== active.id) {
			if (hasVisibleContent) result.push(tag(active, 'end'));
			hasVisibleContent = false;
		}
		active = control;
		for (const part of adapt(run)) {
			if (control && !part.hidden) {
				if (!hasVisibleContent) result.push(tag(control, 'start'));
				hasVisibleContent = true;
			}
			result.push(part);
		}
	}
	if (active && hasVisibleContent) result.push(tag(active, 'end'));
	return result;
}
