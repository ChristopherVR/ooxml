import type { MarkSpec } from 'prosemirror-model';

import { authorColor } from '../review-colors';

export { authorColor };

/** The move name from a revision mark's `move` attr (JSON), for display and linkage. */
export function moveName(value: unknown): string | undefined {
	if (typeof value !== 'string') return undefined;
	try {
		const name = (JSON.parse(value) as { name?: unknown }).name;
		return typeof name === 'string' ? name : undefined;
	} catch {
		return undefined;
	}
}

/** Marks for tracked-change insertions/deletions and comment-range anchors; merged into schema.ts. */
export const reviewMarks = {
	inlineCommentAnchors: {
		inclusive: false,
		attrs: { ids: { default: [] } },
		parseDOM: [
			{
				tag: 'span[data-inline-comment-ids]',
				getAttrs: (el) => ({
					ids: ((el as HTMLElement).dataset.inlineCommentIds ?? '').split(',').filter(Boolean),
				}),
			},
		],
		toDOM: (mark) => [
			'span',
			{
				'data-inline-comment-ids': (mark.attrs.ids as string[]).join(','),
				...((mark.attrs.ids as string[]).length
					? {
							class: 'dve-comment-range',
							'data-comment-ids': (mark.attrs.ids as string[]).join(','),
						}
					: {}),
			},
			0,
		],
	},
	insertion: {
		// `move`: JSON `{ name, rangeId? }` when this is one side of a tracked move (moveTo/moveFrom).
		attrs: {
			author: { default: '' },
			date: { default: null },
			dateUtc: { default: null },
			id: { default: '' },
			move: { default: null },
		},
		parseDOM: [
			{
				tag: 'ins[data-revision-id]',
				getAttrs: (el) => ({
					author: (el as HTMLElement).dataset.author || '',
					date: (el as HTMLElement).dataset.date || null,
					dateUtc: (el as HTMLElement).dataset.dateUtc || null,
					id: (el as HTMLElement).dataset.revisionId || '',
					move: null,
				}),
			},
		],
		toDOM: (mark) => [
			'ins',
			{
				'data-revision-id': mark.attrs.id,
				'data-author': mark.attrs.author,
				...(mark.attrs.date ? { 'data-date': mark.attrs.date } : {}),
				...(mark.attrs.dateUtc ? { 'data-date-utc': mark.attrs.dateUtc } : {}),
				class: mark.attrs.move ? 'dve-revision-insert dve-revision-move' : 'dve-revision-insert',
				...(mark.attrs.move ? { 'data-move': moveName(mark.attrs.move) } : {}),
				style: `--dve-revision-color:${authorColor(mark.attrs.author)}`,
			},
			0,
		],
	},
	deletion: {
		// `move`: JSON `{ name, rangeId? }` when this is one side of a tracked move (moveTo/moveFrom).
		attrs: {
			author: { default: '' },
			date: { default: null },
			dateUtc: { default: null },
			id: { default: '' },
			move: { default: null },
		},
		parseDOM: [
			{
				tag: 'del[data-revision-id]',
				getAttrs: (el) => ({
					author: (el as HTMLElement).dataset.author || '',
					date: (el as HTMLElement).dataset.date || null,
					dateUtc: (el as HTMLElement).dataset.dateUtc || null,
					id: (el as HTMLElement).dataset.revisionId || '',
					move: null,
				}),
			},
		],
		toDOM: (mark) => [
			'del',
			{
				'data-revision-id': mark.attrs.id,
				'data-author': mark.attrs.author,
				...(mark.attrs.date ? { 'data-date': mark.attrs.date } : {}),
				...(mark.attrs.dateUtc ? { 'data-date-utc': mark.attrs.dateUtc } : {}),
				class: mark.attrs.move ? 'dve-revision-delete dve-revision-move' : 'dve-revision-delete',
				...(mark.attrs.move ? { 'data-move': moveName(mark.attrs.move) } : {}),
				style: `--dve-revision-color:${authorColor(mark.attrs.author)}`,
			},
			0,
		],
	},
	comment: {
		// Independent overlapping marks let Yjs merge coauthors' comment anchors.
		excludes: '',
		attrs: { ids: { default: [] } },
		parseDOM: [
			{
				tag: 'span[data-comment-ids]',
				getAttrs: (el) => ({
					ids: ((el as HTMLElement).dataset.commentIds || '').split(',').filter(Boolean),
				}),
			},
		],
		toDOM: (mark) => [
			'span',
			{ class: 'dve-comment-range', 'data-comment-ids': (mark.attrs.ids as string[]).join(',') },
			0,
		],
	},
} satisfies Record<string, MarkSpec>;
