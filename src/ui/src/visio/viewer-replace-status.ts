import type { VisioTextOccurrence } from 'ooxml-core/visio/ui';
import type { ViewerState } from './controller';
import type { FindBar } from './viewer-search';

export const sameTextOccurrence = (a: VisioTextOccurrence, b: VisioTextOccurrence) =>
	a.pageId === b.pageId && a.shapeId === b.shapeId && a.start === b.start && a.end === b.end;

/** Occurrence status is presentation only; the controller and core own all matching. */
export function renderReplaceStatus(
	bar: FindBar,
	state: ViewerState,
	matches: readonly VisioTextOccurrence[],
	current: VisioTextOccurrence | undefined,
	pending: boolean,
	summary: string,
): void {
	const disabled =
		state.loading || state.edit.busy || !state.edit.sourceAvailable || !state.document;
	bar.toggleAttribute('disabled', disabled);
	const unavailable = disabled || !bar.value || !matches.length || pending;
	bar.replaceDisabled = unavailable;
	bar.replaceAllDisabled = unavailable;
	bar.toggleAttribute('navigation-disabled', unavailable);
	const index = current ? matches.findIndex((item) => sameTextOccurrence(item, current)) : -1;
	const occurrences = `${matches.length} occurrence${matches.length === 1 ? '' : 's'}`;
	bar.status = !state.edit.sourceAvailable
		? 'Open a .vsdx file to replace text'
		: state.loading
			? 'Opening diagram'
			: state.edit.busy
				? 'Updating diagram'
				: summary ||
					(!bar.value
						? 'Literal, case-sensitive occurrences'
						: index < 0
							? occurrences
							: `${index + 1} of ${occurrences}`);
	bar.statusTitle = current
		? `Page ${current.pageId}, shape ${current.shapeId}, characters ${current.start + 1} to ${current.end}. Literal case-sensitive matching.`
		: 'Literal case-sensitive matching. Selected groups include their subtree.';
}
