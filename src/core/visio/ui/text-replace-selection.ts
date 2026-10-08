import type { VisioDocument } from '../model';
import type { VisioShapeSelection } from './contract';
import { replaceEntries } from './text-replace-snapshot';
import { replaceIdentity, type VisioTextOccurrence } from './text-replace-types';

/** Own a bounded navigation target. The controller still applies visibility and intent guards. */
export function visioTextOccurrenceSelection(
	model: VisioDocument,
	occurrence: VisioTextOccurrence,
): { readonly pageIndex: number; readonly selection: VisioShapeSelection } | undefined {
	const pageId = replaceIdentity(occurrence.pageId),
		shapeId = replaceIdentity(occurrence.shapeId),
		start = occurrence.start,
		end = occurrence.end;
	if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 0 || end <= start)
		return undefined;
	const entries = replaceEntries(
		model,
		{ pageId, scope: 'all-pages', matchCase: true, query: '' },
		false,
	);
	const entry = entries.find((item) => item.pageId === pageId && item.shapeId === shapeId);
	if (!entry || end > entry.text.length) return undefined;
	return Object.freeze({
		pageIndex: entry.pageIndex,
		selection: Object.freeze({ id: shapeId, name: entry.name, pageId }),
	});
}
