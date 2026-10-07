import { isValidId } from '../collab/validation.js';
import { MAX_MESSAGE_CHARS, MAX_ATTACHMENTS } from './model.js';
import type { StorageLike, UploadableFile } from './store.js';

export interface DraftContext {
	channelId: string;
	threadId?: string;
	editId?: string;
	replyTo?: string;
	/** Separate recovery when a failed send would overwrite a newer compose draft. */
	draftId?: string;
}
export interface ChatDraft {
	text: string;
	files: (UploadableFile & Blob)[];
	/** Names restored without their bytes. Sending requires reattaching or discarding them. */
	missingFiles: string[];
}
export interface SavedDraft extends ChatDraft {
	context: DraftContext;
	updatedAt: number;
}
const idOf = (context: DraftContext) =>
	JSON.stringify([
		context.channelId,
		context.threadId ?? '',
		context.editId ?? '',
		context.draftId ?? '',
	]);
const valid = (context: DraftContext) =>
	isValidId(context.channelId, 64) &&
	[context.threadId, context.editId, context.replyTo, context.draftId].every(
		(id) => id === undefined || isValidId(id),
	);
const names = (input: unknown): string[] =>
	Array.isArray(input)
		? input
				.filter(
					(name): name is string =>
						typeof name === 'string' && name.length > 0 && name.length <= 255,
				)
				.slice(0, MAX_ATTACHMENTS)
		: [];
const empty = (): ChatDraft => ({ text: '', files: [], missingFiles: [] });
const copy = (draft: ChatDraft): ChatDraft => ({
	text: draft.text,
	files: [...draft.files],
	missingFiles: [...draft.missingFiles],
});

/** Device-local text plus session-local binary attachments. No draft enters the shared CRDT. */
export function createDraftStore(storage: StorageLike | undefined, key: string) {
	const rows = new Map<string, SavedDraft>();
	const revisions = new Map<string, number>();
	try {
		const raw = storage?.getItem(key) ?? '[]';
		const parsed: unknown = raw.length <= 524_288 ? JSON.parse(raw) : [];
		if (Array.isArray(parsed))
			for (const row of parsed.slice(0, 500)) {
				if (
					!row ||
					typeof row !== 'object' ||
					!row.context ||
					!valid(row.context) ||
					typeof row.text !== 'string'
				)
					continue;
				const context: DraftContext = {
					channelId: row.context.channelId,
					...(row.context.threadId ? { threadId: row.context.threadId } : {}),
					...(row.context.editId ? { editId: row.context.editId } : {}),
					...(row.context.replyTo ? { replyTo: row.context.replyTo } : {}),
					...(row.context.draftId ? { draftId: row.context.draftId } : {}),
				};
				rows.set(idOf(context), {
					context,
					text: row.text.slice(0, MAX_MESSAGE_CHARS),
					files: [],
					missingFiles: names(row.fileNames),
					updatedAt: Number.isFinite(row.updatedAt) ? row.updatedAt : 0,
				});
			}
	} catch {
		/* Invalid or blocked storage does not affect shared content. */
	}
	const persist = () => {
		try {
			const value = JSON.stringify(
				[...rows.values()].map((row) => ({
					context: row.context,
					text: row.text,
					fileNames: [...row.files.map((file) => file.name), ...row.missingFiles],
					updatedAt: row.updatedAt,
				})),
			);
			if (value.length <= 524_288) storage?.setItem(key, value);
		} catch {
			/* Session drafts remain usable. */
		}
	};
	const set = (context: DraftContext, draft: ChatDraft): number => {
		if (!valid(context)) return -1;
		const id = idOf(context);
		const revision = (revisions.get(id) ?? 0) + 1;
		revisions.set(id, revision);
		if (draft.text || draft.files.length || draft.missingFiles.length)
			rows.set(id, {
				context: { ...context },
				text: draft.text.slice(0, MAX_MESSAGE_CHARS),
				files: draft.files.slice(0, MAX_ATTACHMENTS),
				missingFiles: names(draft.missingFiles),
				updatedAt: Date.now(),
			});
		else rows.delete(id);
		persist();
		return revision;
	};
	return {
		read(context: DraftContext, initialText = ''): ChatDraft {
			const row = rows.get(idOf(context));
			return row ? copy(row) : { ...empty(), text: initialText };
		},
		set,
		clear: (context: DraftContext) => set(context, empty()),
		restore(context: DraftContext, draft: ChatDraft, revision: number) {
			if (revisions.get(idOf(context)) !== revision) return false;
			set(context, draft);
			return true;
		},
		list(): SavedDraft[] {
			return [...rows.values()]
				.sort((a, b) => b.updatedAt - a.updatedAt)
				.map((row) => ({ ...copy(row), context: { ...row.context }, updatedAt: row.updatedAt }));
		},
	};
}
