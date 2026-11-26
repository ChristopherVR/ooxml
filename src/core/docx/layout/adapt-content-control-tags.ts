import type { ContentControl, TextRun } from '../model';
import { definedProps } from '../defined-props';
import type { LayoutRun } from './input';
import { DEFAULT_FONT_SIZE_PT } from './units';
import type { BlockContentControlTags } from './block-content-control-tags';

function tag(control: ContentControl, edge: 'start' | 'end', adjacent?: LayoutRun): LayoutRun {
	return {
		text: edge === 'start' ? (control.alias ?? '').replace(/\s+/g, ' ').trim() : '',
		contentControlTag: edge,
		synthetic: true,
		sourceLength: 0,
		fontSizePt: (adjacent?.fontSizePt ?? DEFAULT_FONT_SIZE_PT) * 0.75,
		...definedProps({ fontFamily: adjacent?.fontFamily }),
	};
}

/** Adds the enclosing controls' tags outside any inline tags already in the paragraph. */
export function adaptBlockContentControlTags(
	runs: LayoutRun[],
	edges: BlockContentControlTags,
): LayoutRun[] {
	let first: LayoutRun | undefined;
	let last: LayoutRun | undefined;
	for (const run of runs) {
		if (run.hidden || run.contentControlTag) continue;
		first ??= run;
		last = run;
	}
	return [
		...edges.start.map((control) => tag(control, 'start', first)),
		...runs,
		...edges.end.map((control) => tag(control, 'end', last)),
	];
}

/** A span can include several source runs and each can produce several font runs. */
export function adaptContentControlTags(
	runs: readonly TextRun[],
	adapt: (run: TextRun) => LayoutRun[],
): LayoutRun[] {
	const result: LayoutRun[] = [];
	let active: ContentControl | undefined;
	let lastVisible: LayoutRun | undefined;
	for (const run of runs) {
		const control = run.contentControl?.appearance === 'tags' ? run.contentControl : undefined;
		if (active && control?.id !== active.id) {
			if (lastVisible) result.push(tag(active, 'end', lastVisible));
			lastVisible = undefined;
		}
		active = control;
		for (const part of adapt(run)) {
			if (control && !part.hidden) {
				if (!lastVisible) result.push(tag(control, 'start', part));
				lastVisible = part;
			}
			result.push(part);
		}
	}
	if (active && lastVisible) result.push(tag(active, 'end', lastVisible));
	return result;
}
