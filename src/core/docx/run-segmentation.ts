// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Whether an edited paragraph can be rebuilt without losing run properties the model does not
// hold. A run read with such properties keeps them as its source XML (`sourceRunPropertiesXml`,
// or `restoredRunPropertiesXml` after a rejected format change): the editor carries it on the text
// through edits, and the writer rebuilds that run's properties from it. Any other run is written
// over the source run element at its own place (inline-source.ts) and keeps that element's
// unmodeled properties, so a change of run boundaries can move them to the wrong text.
import { diffText } from '../text/text-diff';
import { positionText } from './inline-marker-positions';
import type { InlineSlot } from './inline-slots';
import type { Paragraph, TextRun } from './model';
import { samePictureSource } from './inline-source';
import { runHasUnknownProperties } from './write-run-validation';

/** The source XML a run's properties are rebuilt from, if it has one. */
const basisOf = (run: TextRun): string | undefined =>
	run.restoredRunPropertiesXml ?? run.sourceRunPropertiesXml;

/**
 * Whether every character the edit left as it was and that had properties the model does not
 * hold is still in a run with the same source XML, character for character through a diff of the
 * paragraph's text.
 */
function keepsUnmodeledProperties(before: readonly TextRun[], after: readonly TextRun[]): boolean {
	const owners = (runs: readonly TextRun[]) =>
		runs.flatMap((run) => Array.from({ length: positionText(run).length }, () => basisOf(run)));
	const was = owners(before);
	const now = owners(after);
	const kept = (from: number, to: number, shift: number) => {
		for (let at = from; at < to; at++)
			if (was[at] !== undefined && was[at] !== now[at + shift]) return false;
		return true;
	};
	let fromAt = 0;
	let toAt = 0;
	for (const change of diffText(
		before.map(positionText).join(''),
		after.map(positionText).join(''),
	)) {
		if (!kept(fromAt, change.fromStart, toAt - fromAt)) return false;
		fromAt = change.fromEnd;
		toAt = change.toEnd;
	}
	return kept(fromAt, was.length, toAt - fromAt);
}

/**
 * Checks that the edit keeps the run properties the model does not hold, and returns the runs to
 * write on a new run element instead of over the source element at their place: runs without
 * source XML of their own where that element has such properties, once the run boundaries changed
 * (the editor joins equal runs, and text typed at a run's end can be a run of its own). Throws when
 * a picture's or another run's unmodeled properties could be dropped or given to other text.
 */
export function runsToWriteAnew(
	paragraph: Paragraph,
	base: Paragraph | undefined,
	slots: readonly InlineSlot[],
): ReadonlySet<number> {
	const none = new Set<number>();
	const oldRuns = slots.flatMap((slot) => (slot.element ? [slot.element] : []));
	if (!base || !oldRuns.some(runHasUnknownProperties)) return none;
	// Pictures are matched by their media source, so splitting surrounding text does not move
	// their opaque properties to a different run. Retained opaque text still needs its own basis.
	if (
		oldRuns.length === base.runs.length &&
		oldRuns.every((node, index) => {
			if (!runHasUnknownProperties(node)) return true;
			const image = base.runs[index]?.image;
			if (!image) return false;
			// A drawing without a media part (a shape, text box or chart) is the same drawing when
			// its model is unchanged; a changed one is rebuilt from its own source element.
			const remaining = paragraph.runs.filter((run) =>
				image.partName
					? samePictureSource(run.image, image)
					: run.image !== undefined && JSON.stringify(run.image) === JSON.stringify(image),
			);
			return (
				!remaining.length ||
				remaining.every((run) => run.restoredRunPropertiesXml || run.sourceRunPropertiesXml)
			);
		})
	)
		return none;
	// Imported opaque properties or a rejected format revision supply a complete XML basis.
	// The writer imports that basis before applying modeled edits, independently of source slots.
	if (
		paragraph.runs.every(
			(run) =>
				!run.image && !run.equation && (run.restoredRunPropertiesXml || run.sourceRunPropertiesXml),
		)
	)
		return none;
	const sameText =
		paragraph.runs.map((run) => run.text).join('') === base.runs.map((run) => run.text).join('');
	const sameBoundaries =
		paragraph.runs.length === base.runs.length &&
		paragraph.runs.every((run, index) => run.text === base.runs[index]?.text);
	if (sameBoundaries || (!sameText && paragraph.runs.length === base.runs.length)) return none;
	// The boundaries changed. Text keeps its unmodeled properties when each unchanged character
	// that had them is in a run with the same source XML; a run without source XML of its own is
	// then written on a new element, so that it takes none of the element's at its place.
	const object = (run: TextRun) => Boolean(run.image || run.equation);
	if (
		!paragraph.runs.some(object) &&
		!base.runs.some(object) &&
		keepsUnmodeledProperties(base.runs, paragraph.runs)
	)
		return new Set(
			paragraph.runs.flatMap((run, index) => {
				const element = slots[index]?.element;
				return !basisOf(run) && element && runHasUnknownProperties(element) ? [index] : [];
			}),
		);
	throw new Error(
		`Cannot edit paragraph ${paragraph.id}: changing run boundaries could drop unsupported run properties. The original DOCX package remains unchanged.`,
	);
}
