// The list-shaped parts of a shared sheet (comments, conditional formats, validations, links,
// tables, charts, ink, PivotTables) between the model and their shared maps. See part-specs.ts for
// the mapping.
import { formatAddress } from '../address';
import type { Comment, DrawingObject, InkObject, Worksheet } from '../model';
import { type PartIds, type PartSpec, type PrefixStore, readKeyed, writeKeyed } from './keyed';
import {
	chartSpec,
	commentSpec,
	conditionalSpec,
	inkSpec,
	linkSpec,
	replySpec,
	tableSpec,
	validationSpec,
} from './part-specs';
import { pivotShareable } from '../pivot/share-limit';
import { pivotSpec } from './pivot-spec';
import type { ListPart } from './schema';

/** What the sheet's keyed parts need from the binding: item identity and key minting. */
export interface PartHost {
	store: (part: ListPart) => PrefixStore;
	ids: (part: ListPart) => PartIds;
	mint: () => string;
	/** Told about each value a peer wrote that reading left out. */
	refuse?: (part: ListPart, field: string) => void;
}

type Reply = NonNullable<Comment['replies']>[number];

const repliesOwners = (comments: readonly Comment[]): Map<object, string> => {
	const owner = new Map<object, string>();
	for (const comment of comments)
		for (const reply of comment.replies ?? []) owner.set(reply, formatAddress(comment.address));
	return owner;
};

const allReplies = (comments: readonly Comment[]): Reply[] =>
	comments.flatMap((comment) => comment.replies ?? []);

const inksOf = (sheet: Worksheet): InkObject[] =>
	sheet.drawings.filter((d): d is InkObject => d.kind === 'ink');

/**
 * Writes the object lists of `sheet`; `charts` (charts, ink and PivotTables) is skipped for
 * cell-range edits, which cannot touch any of them.
 */
export function writeSheetParts(
	sheet: Worksheet,
	host: PartHost,
	options: { charts: boolean },
): void {
	const put = <T extends object>(part: ListPart, items: readonly T[], spec: PartSpec<T>): void =>
		writeKeyed(host.store(part), items, host.ids(part), spec, host.mint);
	put('comments', sheet.comments, commentSpec);
	put('replies', allReplies(sheet.comments), replySpec(repliesOwners(sheet.comments), new Map()));
	put('cf', sheet.conditionalFormats, conditionalSpec);
	put('dv', sheet.dataValidations, validationSpec);
	put('links', sheet.hyperlinks, linkSpec);
	put('tables', sheet.tables, tableSpec);
	if (options.charts)
		put(
			'charts',
			sheet.drawings.filter(
				(d): d is Extract<DrawingObject, { kind: 'chart' }> => d.kind === 'chart',
			),
			chartSpec,
		);
	if (options.charts) put('ink', inksOf(sheet), inkSpec);
	if (options.charts) put('pivots', sheet.pivotTables ?? [], pivotSpec);
}

/**
 * Reads the object lists named in `parts` (all when undefined) into `sheet`, in place.  Returns whether
 * anything in the sheet changed. Ink is read only from a room whose writers share it (`ink`).
 */
export function applySheetParts(
	sheet: Worksheet,
	host: PartHost,
	parts?: ReadonlySet<string>,
	options: { ink: boolean } = { ink: true },
): boolean {
	const wants = (part: ListPart): PrefixStore | undefined =>
		(part !== 'ink' || options.ink) && (!parts || parts.has(part)) ? host.store(part) : undefined;
	let changed = false;
	const refuse = (part: ListPart) => (field: string) => host.refuse?.(part, field);
	const swap = <T extends object>(
		part: ListPart,
		list: T[],
		spec: PartSpec<T>,
		rebuild: (next: T[]) => void,
	): void => {
		const shared = wants(part);
		const next = shared && readKeyed(shared, list, host.ids(part), spec, refuse(part));
		if (!next) return;
		rebuild(next);
		changed = true;
	};
	const inPlace =
		<T extends object>(list: T[]) =>
		(next: T[]) =>
			void list.splice(0, list.length, ...next);

	const commentsShared = wants('comments');
	const repliesShared = wants('replies');
	if (commentsShared || repliesShared) {
		const owners = repliesOwners(sheet.comments);
		const local = allReplies(sheet.comments);
		let commentsNext: Comment[] | undefined;
		if (commentsShared) {
			commentsNext = readKeyed(
				commentsShared,
				sheet.comments,
				host.ids('comments'),
				commentSpec,
				refuse('comments'),
			);
			if (commentsNext) sheet.comments.splice(0, sheet.comments.length, ...commentsNext);
		}
		const decoded = new Map<object, string>();
		const repliesNext = repliesShared
			? readKeyed(
					repliesShared,
					local,
					host.ids('replies'),
					replySpec(owners, decoded),
					refuse('replies'),
				)
			: undefined;
		if (commentsNext || repliesNext) {
			const grouped = new Map<string, Reply[]>();
			for (const reply of repliesNext ?? local) {
				const at = decoded.get(reply) ?? owners.get(reply);
				if (at === undefined) continue;
				const list = grouped.get(at);
				if (list) list.push(reply);
				else grouped.set(at, [reply]);
			}
			const same = (x: Reply[] | undefined, y: Reply[] | undefined): boolean =>
				(x?.length ?? 0) === (y?.length ?? 0) && (x ?? []).every((r, n) => r === y?.[n]);
			sheet.comments.forEach((comment) => {
				const replies = grouped.get(formatAddress(comment.address));
				if (same(replies, comment.replies)) return;
				if (replies) comment.replies = replies;
				else delete comment.replies;
			});
			changed = true;
		}
	}
	swap('cf', sheet.conditionalFormats, conditionalSpec, inPlace(sheet.conditionalFormats));
	swap('dv', sheet.dataValidations, validationSpec, inPlace(sheet.dataValidations));
	swap('links', sheet.hyperlinks, linkSpec, inPlace(sheet.hyperlinks));
	swap('tables', sheet.tables, tableSpec, (next) => {
		const used = new Set<number>();
		let max = next.reduce((m, t) => Math.max(m, t.id), 0);
		// Concurrent adds can pick the same id; every peer renumbers the later one the same way.
		for (const table of next) {
			if (used.has(table.id)) table.id = ++max;
			used.add(table.id);
		}
		inPlace(sheet.tables)(next);
	});
	// Charts and ink keep their slots among the other drawings, in the shared order; extras go last.
	const slots =
		(kind: 'chart' | 'ink') =>
		(next: DrawingObject[]): void => {
			let k = 0;
			const merged: DrawingObject[] = [];
			for (const drawing of sheet.drawings) {
				if (drawing.kind !== kind) merged.push(drawing);
				else if (k < next.length) merged.push(next[k++] as DrawingObject);
			}
			while (k < next.length) merged.push(next[k++] as DrawingObject);
			sheet.drawings.splice(0, sheet.drawings.length, ...merged);
		};
	const charts = sheet.drawings.filter(
		(d): d is Extract<DrawingObject, { kind: 'chart' }> => d.kind === 'chart',
	);
	swap('charts', charts, chartSpec, slots('chart'));
	swap('ink', inksOf(sheet), inkSpec, slots('ink'));
	// A PivotTable too large to share stays on this peer; only the shared ones are reconciled.
	const pivots = sheet.pivotTables ?? [];
	const unshared = pivots.filter((pivot) => !pivotShareable(pivot));
	swap(
		'pivots',
		pivots.filter((pivot) => pivotShareable(pivot)),
		pivotSpec,
		(next) => {
			sheet.pivotTables = [...next, ...unshared];
		},
	);
	return changed;
}
