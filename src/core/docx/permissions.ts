// Editable ranges (`w:permStart` .. `w:permEnd`), the exceptions of a document protected for
// reading, comments or tracked changes only. The markers stay in the source and are carried
// through edits (inline-markers.ts); this reads which text they cover, clipped to each paragraph,
// so the editor can allow edits there. A range may run across paragraphs and table cells.
import { collectInlineSlots } from './inline-slots';
import { runPositions } from './inline-marker-positions';
import type { Block, Paragraph } from './model';
import type { PermissionRange } from './form-field-model';
import { forEachParagraph } from './parse-warnings';
import { getW, named, type XmlElement } from './xml';

interface MarkerEvent {
	start: boolean;
	id: string;
	position: number;
	group?: string;
	user?: string;
}

/** The markers of paragraphs just parsed, until `resolvePermissions` has joined them. */
const pending = new WeakMap<Paragraph, MarkerEvent[]>();

/** Remembers the editing permission markers of a parsed paragraph, at their character positions. */
export function readParagraphPermissionMarkers(node: XmlElement, paragraph: Paragraph): void {
	if (
		!node.getElementsByTagNameNS('*', 'permStart').length &&
		!node.getElementsByTagNameNS('*', 'permEnd').length
	)
		return;
	const found = collectInlineSlots(node);
	// A paragraph whose runs cannot be matched one for one gives no ranges: nothing is editable.
	if (!found || found.slots.length !== paragraph.runs.length) return;
	const positions = runPositions(paragraph.runs);
	const events: MarkerEvent[] = [];
	for (const { element, boundary } of found.markers) {
		const start = named(element, 'permStart');
		if (!start && !named(element, 'permEnd')) continue;
		const id = getW(element, 'id');
		if (id === undefined) continue;
		const group = getW(element, 'edGrp');
		const user = getW(element, 'ed');
		events.push({
			start,
			id,
			position: positions[boundary] ?? 0,
			...(group !== undefined && { group }),
			...(user !== undefined && { user }),
		});
	}
	if (events.length) pending.set(paragraph, events);
}

type Who = { group?: string; user?: string };

/**
 * Joins the markers read from the paragraphs of `blocks` (in document order, table cells
 * included) into each paragraph's `permissions`. A start without an end in its paragraph runs to
 * the end of the paragraph and on through the paragraphs after it, up to the end marker.
 */
export function resolvePermissions(blocks: Block[]): void {
	const open = new Map<string, Who>();
	forEachParagraph(blocks, (paragraph) => {
		const events = pending.get(paragraph) ?? [];
		const end = runPositions(paragraph.runs).at(-1) ?? 0;
		const ranges: PermissionRange[] = [];
		type Begin = Who & { from: number; previous?: true };
		const begun = new Map<string, Begin>(
			[...open].map(([id, who]) => [id, { ...who, from: 0, previous: true }]),
		);
		const add = (id: string, begin: Begin, to: number, next?: true) =>
			ranges.push({
				id,
				...(begin.group !== undefined && { group: begin.group }),
				...(begin.user !== undefined && { user: begin.user }),
				from: begin.from,
				to,
				...(begin.previous && { fromPrevious: true as const }),
				...(next && { toNext: true as const }),
			});
		for (const event of events) {
			if (event.start) {
				const who: Who = {
					...(event.group !== undefined && { group: event.group }),
					...(event.user !== undefined && { user: event.user }),
				};
				begun.set(event.id, { ...who, from: event.position });
				open.set(event.id, who);
				continue;
			}
			const begin = begun.get(event.id);
			if (begin) add(event.id, begin, event.position);
			begun.delete(event.id);
			open.delete(event.id);
		}
		// Ranges still open at the end of the paragraph cover the rest of it.
		for (const [id, begin] of begun) add(id, begin, end, true);
		if (ranges.length) paragraph.permissions = ranges;
		pending.delete(paragraph);
	});
}
