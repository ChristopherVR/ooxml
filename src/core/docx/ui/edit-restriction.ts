// Enforces Review > Restrict Editing (`w:documentProtection w:edit="readOnly|comments|trackedChanges"`)
// like Word: read only allows edits only inside the editable ranges (`w:permStart`/`w:permEnd` for
// everyone or the current user), comments only adds the comments as well, and tracked changes
// forces Track Changes on and keeps it from being turned off. Forms protection has its own guard
// (form-field-guard.ts).
import type { Node } from 'prosemirror-model';
import { Plugin, PluginKey, type EditorState, type Transaction } from 'prosemirror-state';
import {
	AddMarkStep,
	AttrStep,
	DocAttrStep,
	RemoveMarkStep,
	ReplaceAroundStep,
	ReplaceStep,
} from 'prosemirror-transform';
import type { DocumentProtection, PermissionRange } from '../model';

export type EditRestriction = 'readOnly' | 'comments' | 'trackedChanges';

/** The enforced restriction this guard applies, if any (forms protection is not one of them). */
export function editRestrictionOf(
	protection: DocumentProtection | undefined,
): EditRestriction | undefined {
	if (!protection?.enforced) return undefined;
	const { edit } = protection;
	return edit === 'readOnly' || edit === 'comments' || edit === 'trackedChanges' ? edit : undefined;
}

/** Whether Track Changes is forced on and cannot be turned off. */
export const forcesTrackChanges = (protection: DocumentProtection | undefined): boolean =>
	editRestrictionOf(protection) === 'trackedChanges';

/** Whether headers, footers and notes are closed to editing (anything but tracked changes). */
export function locksStories(protection: DocumentProtection | undefined): boolean {
	return (
		Boolean(protection?.enforced) &&
		(protection?.edit === 'forms' || editRestrictionOf(protection) !== undefined) &&
		protection?.edit !== 'trackedChanges'
	);
}

/** An editable range in editor positions. */
export interface EditableRange {
	id?: string;
	/** The range continues from the previous paragraph, or into the next one. */
	fromPrevious?: true;
	toNext?: true;
	from: number;
	to: number;
	group?: string;
	user?: string;
}

export const editRestrictionKey = new PluginKey<EditableRange[]>('docx-edit-restriction');

const isRemote = (tr: Transaction): boolean => tr.getMeta('y-sync$') !== undefined;

/** The ranges the paragraphs' `permissions` cover, in editor positions. */
export function permissionRanges(doc: Node): EditableRange[] {
	const ranges: EditableRange[] = [];
	doc.descendants((node, pos) => {
		if (node.type.name === 'paragraph') {
			const permissions = node.attrs.permissions as PermissionRange[] | null | undefined;
			const end = pos + node.nodeSize - 1;
			for (const item of permissions ?? [])
				ranges.push({
					id: item.id,
					...(item.fromPrevious && { fromPrevious: true as const }),
					...(item.toNext && { toNext: true as const }),
					from: Math.min(pos + 1 + item.from, end),
					to: Math.min(pos + 1 + item.to, end),
					...(item.group !== undefined && { group: item.group }),
					...(item.user !== undefined && { user: item.user }),
				});
			return false;
		}
		return true;
	});
	return ranges;
}

/** Whether `user` may edit the range: Word's "Everyone" exception, or the range's own user. */
export const rangeAllows = (range: EditableRange, user: string): boolean =>
	range.group === 'everyone' ||
	(range.user !== undefined && range.user.toLowerCase() === user.toLowerCase());

/** The ranges `user` may edit in `state`, mapped through every edit made since the document opened. */
export const editableRanges = (state: EditorState, user: string): EditableRange[] =>
	(editRestrictionKey.getState(state) ?? []).filter((range) => rangeAllows(range, user));

/** Whether the transaction only changes the document the way the restriction allows. */
export function transactionAllowed(
	tr: Transaction,
	restriction: EditRestriction,
	ranges: readonly EditableRange[],
): boolean {
	return tr.steps.every((step, index) => {
		if (step instanceof DocAttrStep) {
			const { attr, value } = step as unknown as { attr: string; value: unknown };
			if (restriction === 'trackedChanges') return !(attr === 'trackChanges' && !value);
			return restriction === 'comments' && attr === 'commentThreads';
		}
		// Block ids are bookkeeping the editor assigns.
		if (step instanceof AttrStep && (step as unknown as { attr: string }).attr === 'id')
			return true;
		if (restriction === 'trackedChanges') return true;
		if (
			restriction === 'comments' &&
			(step instanceof AddMarkStep || step instanceof RemoveMarkStep) &&
			(step as unknown as { mark: { type: { name: string } } }).mark.type.name === 'comment'
		)
			return true;
		const map = tr.mapping.slice(0, index);
		const within = (from: number, to: number) =>
			ranges.some((range) => from >= map.map(range.from, -1) && to <= map.map(range.to, 1));
		if (step instanceof AttrStep) {
			const { pos } = step as unknown as { pos: number };
			return within(pos, pos + 1);
		}
		if (
			step instanceof ReplaceStep ||
			step instanceof ReplaceAroundStep ||
			step instanceof AddMarkStep ||
			step instanceof RemoveMarkStep
		) {
			const { from, to } = step as unknown as { from: number; to: number };
			return within(from, to);
		}
		return false;
	});
}

/**
 * Rejects local edits the enforced restriction forbids. `user` names the person whose exceptions
 * count (the review author); the editable ranges follow the text through edits.
 */
export function editRestrictionPlugin(
	protection: () => DocumentProtection | undefined,
	user: () => string,
): Plugin<EditableRange[]> {
	return new Plugin<EditableRange[]>({
		key: editRestrictionKey,
		state: {
			init: (_config, state) => permissionRanges(state.doc),
			apply: (tr, ranges) =>
				tr.getMeta(editRestrictionKey) === 'reset'
					? permissionRanges(tr.doc)
					: tr.docChanged
						? ranges.map((range) => ({
								...range,
								from: tr.mapping.map(range.from, -1),
								to: tr.mapping.map(range.to, 1),
							}))
						: ranges,
		},
		// The paragraphs' `permissions` follow the text through edits, so saving and rebuilding the
		// editor from the model see the ranges where they are now.
		appendTransaction: (transactions, _old, state) => {
			if (!transactions.some((tr) => tr.docChanged && !isRemote(tr))) return null;
			const ranges = editRestrictionKey.getState(state) ?? [];
			if (!ranges.length) return null;
			const tr = state.tr;
			const byParagraph = new Map<number, PermissionRange[]>();
			for (const range of ranges) {
				const $from = state.doc.resolve(Math.min(range.from, state.doc.content.size));
				if ($from.parent.type.name !== 'paragraph') continue;
				const pos = $from.before();
				const base = pos + 1;
				const end = pos + $from.parent.nodeSize - 1;
				byParagraph.set(pos, [
					...(byParagraph.get(pos) ?? []),
					{
						id: range.id ?? '0',
						...(range.group !== undefined && { group: range.group }),
						...(range.user !== undefined && { user: range.user }),
						from: range.from - base,
						to: Math.min(range.to, end) - base,
						...(range.fromPrevious && { fromPrevious: true as const }),
						...(range.toNext && { toNext: true as const }),
					},
				]);
			}
			for (const [pos, list] of byParagraph) {
				const node = state.doc.nodeAt(pos);
				if (node && JSON.stringify(node.attrs.permissions) !== JSON.stringify(list))
					tr.setNodeAttribute(pos, 'permissions', list);
			}
			return tr.docChanged ? tr.setMeta('addToHistory', false) : null;
		},
		filterTransaction: (tr, state) => {
			const restriction = editRestrictionOf(protection());
			if (!restriction || !tr.docChanged || isRemote(tr)) return true;
			return transactionAllowed(tr, restriction, editableRanges(state, user()));
		},
	});
}

/** Why Restrict Editing's exceptions could not change the selection. */
export type ExceptionResult = { tr: Transaction } | { error: 'spansParagraphs' | 'empty' };

const touches = (range: { from: number; to: number }, from: number, to: number) =>
	range.from <= to && range.to >= from;

/**
 * Marks the selection as an editable region for everyone (Restrict Editing > Exceptions), or
 * takes the mark off it. The ranges are kept per paragraph; one that runs on into the paragraph
 * before or after (read from a file) is left to Word.
 */
export function setEditableRegion(
	state: EditorState,
	on: boolean,
	from: number = state.selection.from,
	to: number = state.selection.to,
): ExceptionResult {
	if (from === to) return { error: 'empty' };
	const all = editRestrictionKey.getState(state) ?? permissionRanges(state.doc);
	let next = Math.max(0, ...all.map((range) => Number(range.id) || 0)) + 1;
	const tr = state.tr;
	let blocked = false;
	state.doc.nodesBetween(from, to, (node, pos) => {
		if (node.type.name !== 'paragraph') return true;
		const base = pos + 1;
		const end = pos + node.nodeSize - 1;
		const own = all.filter((range) => range.from >= base && range.to <= end);
		const a = Math.max(from, base) - base;
		const b = Math.min(to, end) - base;
		if (b <= a && !(node.content.size === 0 && on)) return false;
		if (
			own.some(
				(range) => (range.fromPrevious || range.toNext) && touches(range, base + a, base + b),
			)
		) {
			blocked = true;
			return false;
		}
		const everyone = own.filter((range) => range.group === 'everyone');
		const others = own.filter((range) => range.group !== 'everyone');
		const relative = (range: EditableRange): PermissionRange => ({
			id: range.id ?? String(next++),
			...(range.group !== undefined && { group: range.group }),
			...(range.user !== undefined && { user: range.user }),
			from: range.from - base,
			to: range.to - base,
			...(range.fromPrevious && { fromPrevious: true as const }),
			...(range.toNext && { toNext: true as const }),
		});
		let kept: PermissionRange[];
		if (on) {
			const merged = everyone.filter((range) => touches(range, base + a, base + b));
			const rest = everyone.filter((range) => !merged.includes(range));
			const start = Math.min(a, ...merged.map((range) => range.from - base));
			const finish = Math.max(b, ...merged.map((range) => range.to - base));
			kept = [
				...rest.map(relative),
				{ id: merged[0]?.id ?? String(next++), group: 'everyone', from: start, to: finish },
			];
		} else {
			kept = [];
			for (const range of everyone) {
				const r = relative(range);
				if (!touches(range, base + a, base + b) || r.to <= a || r.from >= b) kept.push(r);
				else {
					if (r.from < a) kept.push({ ...r, to: a });
					if (r.to > b) kept.push({ ...r, id: r.from < a ? String(next++) : r.id, from: b });
				}
			}
		}
		kept.push(...others.map(relative));
		kept.sort((x, y) => x.from - y.from);
		tr.setNodeAttribute(pos, 'permissions', kept.length ? kept : null);
		return false;
	});
	if (blocked) return { error: 'spansParagraphs' };
	return { tr: tr.setMeta(editRestrictionKey, 'reset') };
}
