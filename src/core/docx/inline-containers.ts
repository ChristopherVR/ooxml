// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// The hyperlinks, simple fields and content controls a rebuilt paragraph's runs are grouped into,
// and which source container each group continues: the only one with its control or simple field,
// else the one with the same link, control or field that its text moved to (through the character
// diff; every part of a hyperlink that was split), else one with the same link, control or field
// and the same text (it was moved within the paragraph). A group whose link, control or field is
// new takes over only the markers of the container its text came from: first one whose unchanged
// text is exactly where the group is.
import type { ContentControl } from './content-control-model';
import { sameFieldFlags } from './field-lock';
import { moveMarkers, positionText, runPositions } from './inline-marker-positions';
import type { InlineSlot } from './inline-slots';
import { diffText } from '../text/text-diff';
import type { HyperlinkInfo, TextRun } from './model';
import type { XmlElement } from './xml';

export type ContainerGroup =
	| { kind: 'link'; link: HyperlinkInfo }
	| { kind: 'control'; control: ContentControl }
	| {
			kind: 'simpleField';
			instr: string;
			instanceId: string | undefined;
			fieldFlags: TextRun['fieldFlags'];
	  };

function sameLink(a: HyperlinkInfo | undefined, b: HyperlinkInfo | undefined): boolean {
	if (!a && !b) return true;
	if (!a || !b) return false;
	return a.href === b.href && a.anchor === b.anchor && a.tooltip === b.tooltip;
}

/** The wrapper a run is written inside: a hyperlink, or a simple field (`w:fldSimple`). */
export function containerGroup(run: TextRun): ContainerGroup | undefined {
	if (run.contentControl) return { kind: 'control', control: run.contentControl };
	if (run.link) return { kind: 'link', link: run.link };
	if (run.field?.simple)
		return {
			kind: 'simpleField',
			instr: run.field.instr,
			instanceId: run.fieldInstanceId,
			fieldFlags: run.fieldFlags,
		};
	return undefined;
}

export function sameGroup(a: ContainerGroup | undefined, b: ContainerGroup | undefined): boolean {
	if (!a || !b) return a === b;
	if (a.kind === 'control' || b.kind === 'control')
		return a.kind === 'control' && b.kind === 'control' && a.control.id === b.control.id;
	if (a.kind === 'link' && b.kind === 'link') return sameLink(a.link, b.link);
	return (
		a.kind === 'simpleField' &&
		b.kind === 'simpleField' &&
		a.instr === b.instr &&
		a.instanceId === b.instanceId &&
		sameFieldFlags(a.fieldFlags, b.fieldFlags)
	);
}

/** Consecutive runs written into one container: its group and run range. */
interface RunGroup {
	group: ContainerGroup;
	start: number;
	end: number;
}

/** The groups of `runs`; with `elementAt`, runs read from two source elements stay apart. */
function groupsOf(
	runs: readonly TextRun[],
	elementAt?: (index: number) => XmlElement | undefined,
): RunGroup[] {
	const groups: RunGroup[] = [];
	for (const [at, run] of runs.entries()) {
		const group = containerGroup(run);
		const last = groups.at(-1);
		if (!group) continue;
		if (
			last?.end === at &&
			sameGroup(last.group, group) &&
			elementAt?.(at) === elementAt?.(last.start)
		)
			last.end = at + 1;
		else groups.push({ group, start: at, end: at + 1 });
	}
	return groups;
}

/** The source containers a group of the rebuilt paragraph continues. */
export interface ContainerSource {
	containers: ReadonlySet<XmlElement | undefined>;
	/** The same link, control or field, so its container can be reused. */
	same: boolean;
	/** A further part of a split hyperlink: written as a copy of the source element. */
	clone?: boolean;
	/**
	 * A control or field matched by its identity: the markers read in it go with it, mapped through
	 * its own text (`followContainers`).
	 */
	follow?: boolean;
	/** Character positions of the source container's text, before the edit. */
	from: { start: number; end: number };
	/** Character positions of the group's text, after the edit. */
	to: { start: number; end: number };
	/** Indices of the source container's runs and of the group's runs. */
	fromRuns: { start: number; end: number };
	toRuns: { start: number; end: number };
}

/** For each group of `runs`, keyed by its first run's index, the source containers it continues. */
export function containerSources(
	runs: readonly TextRun[],
	base: readonly TextRun[] | undefined,
	slots: readonly InlineSlot[],
): Map<number, ContainerSource> {
	const result = new Map<number, ContainerSource>();
	const targets = groupsOf(runs);
	// Two neighbouring links to one place are two source containers.
	const sources = groupsOf(base ?? [], (at) => slots[at]?.container);
	if (!targets.length || !sources.length) return result;
	const targetAt = runPositions(runs);
	const sourceAt = runPositions(base ?? []);
	// Where each source group's text went.
	const edges = moveMarkers(
		base ?? [],
		runs,
		sources.flatMap((source) => [
			{ position: sourceAt[source.start]!, opening: true, exact: true },
			{ position: sourceAt[source.end]!, opening: false, exact: true },
		]),
	);
	// Where a source group's text went: its edges cross when text was typed into an empty one.
	const movedRange = (index: number) => {
		const [start, end] = [edges[2 * index]!.position, edges[2 * index + 1]!.position];
		return { from: Math.min(start, end), to: Math.max(start, end) };
	};
	const overlap = (target: RunGroup, index: number) => {
		const { from, to } = movedRange(index);
		const [start, end] = [targetAt[target.start]!, targetAt[target.end]!];
		const shared = Math.min(end, to) - Math.max(start, from);
		if (shared > 0) return shared;
		// A container empty before or after the edit meets the group at a point, but a link whose
		// text was all deleted is gone: only its text tells it from a new link to the same place.
		const source = sources[index]!;
		if (
			source.group.kind === 'link' &&
			from === to &&
			sourceAt[source.start]! < sourceAt[source.end]!
		)
			return 0;
		return (from === to || start === end) && from <= end && start <= to ? 0.5 : 0;
	};
	// The text both sides share (the character diff's unchanged stretches), as source ranges with
	// the shift that takes them to the rebuilt paragraph.
	const beforeText = (base ?? []).map(positionText).join('');
	const afterText = runs.map(positionText).join('');
	const unchanged: { start: number; end: number; shift: number }[] = [];
	let fromAt = 0;
	let toAt = 0;
	for (const change of diffText(beforeText, afterText)) {
		if (change.fromStart > fromAt)
			unchanged.push({ start: fromAt, end: change.fromStart, shift: toAt - fromAt });
		fromAt = change.fromEnd;
		toAt = change.toEnd;
	}
	if (beforeText.length > fromAt)
		unchanged.push({ start: fromAt, end: beforeText.length, shift: toAt - fromAt });
	/** How much of a source group's text the group still holds unchanged. */
	const surviving = (target: RunGroup, index: number) => {
		const source = sources[index]!;
		const [start, end] = [targetAt[target.start]!, targetAt[target.end]!];
		let count = 0;
		for (const stretch of unchanged) {
			const from = Math.max(stretch.start, sourceAt[source.start]!) + stretch.shift;
			const to = Math.min(stretch.end, sourceAt[source.end]!) + stretch.shift;
			count += Math.max(0, Math.min(to, end) - Math.max(from, start));
		}
		return count;
	};
	const textOf = (from: readonly TextRun[], group: RunGroup) =>
		from.slice(group.start, group.end).map(positionText).join('');
	const taken = new Set<number>();
	const claim = (
		target: RunGroup,
		index: number,
		same: boolean,
		{ clone = false, identity = false } = {},
	) => {
		const source = sources[index]!;
		taken.add(index);
		const to = { start: targetAt[target.start]!, end: targetAt[target.end]! };
		result.set(target.start, {
			containers: new Set(slots.slice(source.start, source.end).map((slot) => slot?.container)),
			same,
			...(clone && { clone }),
			...(identity && { follow: true }),
			from: { start: sourceAt[source.start]!, end: sourceAt[source.end]! },
			to,
			fromRuns: { start: source.start, end: source.end },
			toRuns: { start: target.start, end: target.end },
		});
	};
	const open = (target: RunGroup) => !result.has(target.start);
	// A control or simple field the paragraph holds once, before and after the edit, keeps its
	// markup wherever its text went: a control has its id, a simple field its instance. A link has
	// nothing but its text to tell it from a new link to the same place, so the passes below match
	// it by its text.
	const like = (groups: readonly RunGroup[], target: RunGroup) =>
		groups.filter((group) => sameGroup(group.group, target.group));
	for (const target of targets) {
		if (target.group.kind === 'link') continue;
		const [source, ...more] = like(sources, target);
		if (source && !more.length && like(targets, target).length === 1)
			claim(target, sources.indexOf(source), true, { identity: true });
	}
	// The same link, control or field that the group's text moved to, widest overlap first.
	const moved = targets
		.flatMap((target) =>
			sources.flatMap((source, index) => {
				// A link that had text continues one holding some of that text unchanged.
				const shared = !sameGroup(source.group, target.group)
					? 0
					: source.group.kind === 'link' && sourceAt[source.start]! < sourceAt[source.end]!
						? surviving(target, index)
						: overlap(target, index);
				return shared > 0 ? [{ target, index, shared }] : [];
			}),
		)
		.sort((a, b) => b.shared - a.shared);
	// Every part of a split hyperlink keeps its markup; a control or field has one id to keep.
	for (const { target, index } of moved)
		if (open(target) && (!taken.has(index) || target.group.kind === 'link'))
			claim(target, index, true, { clone: taken.has(index) });
	// A container whose text did not change and is where a group of its kind with a new link,
	// control or field now is (one with a source of its own continues that source instead).
	const ownSource = (target: RunGroup) =>
		sources.some((other, at) => !taken.has(at) && sameGroup(other.group, target.group));
	for (const [index, source] of sources.entries()) {
		if (taken.has(index)) continue;
		const text = textOf(base ?? [], source);
		const { from, to } = movedRange(index);
		const target = targets.find(
			(candidate) =>
				open(candidate) &&
				!ownSource(candidate) &&
				candidate.group.kind === source.group.kind &&
				targetAt[candidate.start] === from &&
				targetAt[candidate.end] === to &&
				textOf(runs, candidate) === text,
		);
		if (target) claim(target, index, false);
	}
	// The same link, control or field moved elsewhere in the paragraph, which carries its text,
	// unless that text is still where it was (an equal label elsewhere took the link).
	const after = runs.map(positionText).join('');
	const stayed = (index: number) => {
		const { from, to } = movedRange(index);
		return after.slice(from, to) === textOf(base ?? [], sources[index]!);
	};
	for (const target of targets.filter(open)) {
		const text = textOf(runs, target);
		const index = sources.findIndex(
			(source, at) =>
				!taken.has(at) &&
				sameGroup(source.group, target.group) &&
				textOf(base ?? [], source) === text &&
				!stayed(at),
		);
		if (index >= 0) claim(target, index, true);
	}
	// A changed link, control or field keeps the markers of the container its text came from.
	for (const target of targets.filter(open)) {
		let found = -1;
		let best = 0;
		for (const [index, source] of sources.entries()) {
			const shared =
				!taken.has(index) && source.group.kind === target.group.kind ? overlap(target, index) : 0;
			if (shared > best) {
				best = shared;
				found = index;
			}
		}
		if (found >= 0) claim(target, found, false);
	}
	return result;
}
