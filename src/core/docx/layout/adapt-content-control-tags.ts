import type { ContentControl, TextRun } from '../model';
import type { LayoutRun } from './input';
import type { BlockContentControlTags } from './block-content-control-tags';
import { contentControlTagMetrics } from './content-control-tag-metrics';

function tag(control: ContentControl, edge: 'start' | 'end', uncounted = false): LayoutRun {
	const structural = control.level !== 'inline';
	return {
		...(uncounted && { contentControlTagUncounted: true as const }),
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

/** The tagged controls around a run, outermost first. */
const taggedChain = (run: TextRun): ContentControl[] =>
	[...(run.outerContentControls ?? []), ...(run.contentControl ? [run.contentControl] : [])].filter(
		(control) => control.appearance === 'tags',
	);

/** Every control around a run, outermost first, whatever its appearance. */
const fullChain = (run: TextRun): ContentControl[] => [
	...(run.outerContentControls ?? []),
	...(run.contentControl ? [run.contentControl] : []),
];

/** A span can include several source runs and each can produce several font runs. */
export function adaptContentControlTags(
	runs: readonly TextRun[],
	adapt: (run: TextRun) => LayoutRun[],
): LayoutRun[] {
	const result: LayoutRun[] = [];
	const open: { control: ContentControl; visible: boolean; uncounted: boolean }[] = [];
	// Set when an untagged control ended since the last visible text: the tags that follow take
	// no width when Word breaks the line.
	let afterUntagged = false;
	let previousChain: ContentControl[] = [];
	const closeFrom = (keep: number) => {
		while (open.length > keep) {
			const { control, visible } = open.pop()!;
			if (visible) result.push(tag(control, 'end', afterUntagged));
		}
	};
	for (const run of runs) {
		const full = fullChain(run);
		afterUntagged = previousChain.some(
			(control) => control.appearance !== 'tags' && !full.some((other) => other.id === control.id),
		);
		previousChain = full;
		const chain = taggedChain(run);
		let keep = 0;
		while (keep < open.length && keep < chain.length && open[keep]!.control.id === chain[keep]!.id)
			keep++;
		closeFrom(keep);
		for (const control of chain.slice(keep))
			open.push({ control, visible: false, uncounted: afterUntagged });
		for (const part of adapt(run)) {
			if (!part.hidden)
				for (const entry of open)
					if (!entry.visible) {
						entry.visible = true;
						result.push(tag(entry.control, 'start', entry.uncounted));
					}
			result.push(part);
		}
	}
	afterUntagged = previousChain.some((control) => control.appearance !== 'tags');
	closeFrom(0);
	return result;
}
