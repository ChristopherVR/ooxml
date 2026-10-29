// The list-shaped parts of a shared sheet (comments, conditional formats, validations, links,
// tables, charts) between the model and their shared maps. See part-specs.ts for the mapping.
import { formatAddress } from '../address';
import type { Comment, DrawingObject, Worksheet } from '../model';
import { type PartIds, type PartSpec, type PrefixStore, readKeyed, writeKeyed } from './keyed';
import {
	chartSpec,
	commentSpec,
	conditionalSpec,
	linkSpec,
	replySpec,
	tableSpec,
	validationSpec,
} from './part-specs';
import type { ListPart } from './schema';

/** What the sheet's keyed parts need from the binding: item identity and key minting. */
export interface PartHost {
	store: (part: ListPart) => PrefixStore;
	ids: (part: ListPart) => PartIds;
	mint: () => string;
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

/** Writes the object lists of `sheet`; `charts` is skipped for edits that cannot touch drawings. */
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
}

/**
 * Reads the object lists named in `parts` (all when undefined) into `sheet`, in place.  Returns whether
 * anything in the sheet changed.
 */
export function applySheetParts(
	sheet: Worksheet,
	host: PartHost,
	parts?: ReadonlySet<string>,
): boolean {
	const wants = (part: ListPart): PrefixStore | undefined =>
		!parts || parts.has(part) ? host.store(part) : undefined;
	let changed = false;
	const swap = <T extends object>(
		part: ListPart,
		list: T[],
		spec: PartSpec<T>,
		rebuild: (next: T[]) => void,
	): void => {
		const shared = wants(part);
		const next = shared && readKeyed(shared, list, host.ids(part), spec);
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
			commentsNext = readKeyed(commentsShared, sheet.comments, host.ids('comments'), commentSpec);
			if (commentsNext) sheet.comments.splice(0, sheet.comments.length, ...commentsNext);
		}
		const decoded = new Map<object, string>();
		const repliesNext = repliesShared
			? readKeyed(repliesShared, local, host.ids('replies'), replySpec(owners, decoded))
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
	const charts = sheet.drawings.filter(
		(d): d is Extract<DrawingObject, { kind: 'chart' }> => d.kind === 'chart',
	);
	swap('charts', charts, chartSpec, (next) => {
		// Charts keep their slots among the other drawings, in the shared order; extras go last.
		let k = 0;
		const merged: DrawingObject[] = [];
		for (const drawing of sheet.drawings) {
			if (drawing.kind !== 'chart') merged.push(drawing);
			else if (k < next.length) merged.push(next[k++] as DrawingObject);
		}
		while (k < next.length) merged.push(next[k++] as DrawingObject);
		sheet.drawings.splice(0, sheet.drawings.length, ...merged);
	});
	return changed;
}
