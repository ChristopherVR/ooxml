// The rule of `validateTeamsUpdate` for the chat directory (`teams:chats`, one row per chat in the
// workspace document). A row says who is in a chat, so only that chat's people may change it, and
// what kind of chat it is and who created it never change. A group's people live in one field each
// (`memberKey`), so concurrent changes merge and each field can be judged alone. DOM-free.
import * as Y from 'yjs';
import {
	CHATS_MAP,
	type ChatInfo,
	directChatId,
	isChatUserId,
	isGroupChatIdOf,
	memberStanding,
	readChatRow,
	userOfMemberKey,
} from './chat-records';
import { sanitizeChannelName } from './model';
import type { UpdateActor, UpdateViolation } from './validate-update';
import { isSameRow, sameContent } from './validate-owned';

const ROW_FIELDS = new Set(['kind', 'name', 'members', 'createdAt', 'createdBy', 'meetingId']);

/** A field of a chat row to put back: `value` is what it held, undefined when it was absent. */
export interface RowRevert {
	chat: string;
	key: string;
	value: unknown;
}

/** Undo what a verdict's `reverts` list, after its update has been applied to `doc`. */
export function applyUpdateReverts(doc: Y.Doc, reverts: readonly RowRevert[]): void {
	const chats = doc.getMap<unknown>(CHATS_MAP);
	doc.transact(() => {
		for (const { chat, key, value } of reverts) {
			const row = chats.get(chat);
			if (!(row instanceof Y.Map)) continue;
			if (value === undefined) row.delete(key);
			else row.set(key, value);
		}
	});
}

/** The row as a chat, or null when a reader would have to drop any part of it. */
function strictChat(id: string, row: Y.Map<unknown>): ChatInfo | null {
	const chat = readChatRow(id, row);
	const raw = row.get('members');
	if (!chat || !Array.isArray(raw) || raw.length !== chat.members.length) return null;
	return [...row.keys()].every((key) => ROW_FIELDS.has(key)) ? chat : null;
}

const memberIds = (chat: ChatInfo): Set<string> => new Set(chat.members.map((m) => m.id));
const sameIds = (a: Set<string>, b: Set<string>): boolean =>
	a.size === b.size && [...a].every((id) => b.has(id));
const isCleanName = (value: unknown): value is string =>
	typeof value === 'string' && sanitizeChannelName(value) === value;

function checkNewRow(
	id: string,
	row: Y.Map<unknown>,
	actor: UpdateActor,
	at: { map: string; key: string },
): UpdateViolation | null {
	const chat = strictChat(id, row);
	if (!chat) return { code: 'malformed-row', ...at };
	if (row.get('createdBy') !== actor.id) return { code: 'forged-author', ...at };
	if (!memberIds(chat).has(actor.id)) return { code: 'not-member', ...at };
	if (chat.kind === 'group')
		return isGroupChatIdOf(id, actor.id) ? null : { code: 'forged-author', ...at };
	const [a, b] = chat.members;
	return a && b && id === directChatId(a.id, b.id) ? null : { code: 'malformed-row', ...at };
}

function checkDirectEdit(
	id: string,
	prior: Y.Map<unknown>,
	next: Y.Map<unknown>,
	actor: UpdateActor,
	at: { map: string; key: string },
): UpdateViolation | null {
	const was = readChatRow(id, prior);
	if (!was) return { code: 'foreign-row', ...at };
	const had = memberIds(was);
	if (!had.has(actor.id)) return { code: 'not-member', ...at };
	const now = strictChat(id, next);
	if (!now || now.kind !== was.kind) return { code: 'malformed-row', ...at };
	const replaced = !(next._item && prior._item && Y.compareIDs(next._item.id, prior._item.id));
	return replaced && next.get('createdBy') === actor.id && sameIds(had, memberIds(now))
		? null
		: { code: 'foreign-row', ...at };
}

/**
 * Whether the actor may set this field of a group to `value`, or null when no client would write
 * it. A member renames and adds; removing is for yourself, or a moderator who is a member.
 */
function mayWriteGroupField(
	key: string,
	value: unknown,
	actor: UpdateActor,
	isMember: boolean,
): boolean | null {
	if (key === 'name') return isCleanName(value) ? isMember : null;
	const userId = userOfMemberKey(key);
	if (userId === null || !isChatUserId(userId)) return null;
	if (value === null) return userId === actor.id || (isMember && actor.moderator === true);
	return isCleanName(value) && value !== '' ? isMember : null;
}

function checkGroupEdit(
	id: string,
	prior: Y.Map<unknown>,
	next: Y.Map<unknown>,
	actor: UpdateActor,
	at: { map: string; key: string },
	reverts: RowRevert[],
): UpdateViolation | null {
	if (!(next._item && prior._item && Y.compareIDs(next._item.id, prior._item.id)))
		return { code: 'foreign-row', ...at };
	const isMember = memberStanding(prior, actor.id) === 'member';
	const undo: RowRevert[] = [];
	for (const key of new Set([...prior.keys(), ...next.keys()])) {
		const value = prior.get(key);
		const written = next.get(key);
		if (sameContent(value, written)) continue;
		if (key !== 'name' && ROW_FIELDS.has(key)) return { code: 'foreign-row', ...at };
		const may = mayWriteGroupField(key, written, actor, isMember);
		if (may === null) return { code: 'malformed-row', ...at };
		if (!may) undo.push({ chat: id, key, value });
	}
	reverts.push(...undo);
	return null;
}

/**
 * One directory row. A new row names the actor as creator and member, under an id bound to them: a
 * group id they minted, or the 1:1 id of the pair. After that `kind`, `members`, `createdAt`,
 * `createdBy` and `meetingId` never change and a row is never removed. A 1:1 row is immutable,
 * except that either of its two people may replace it with an equal one, because both can open the
 * chat at once.
 *
 * In a group a member may rename it and add people, and removes only themselves unless they are a
 * moderator. Someone who never belonged is refused. Someone who belongs or belonged but may not
 * make this change may simply be out of date (removed while offline, no longer a moderator): the
 * update is accepted and the field goes into `reverts`, because a refusal would repeat on every
 * reconnect and stop their sync for good.
 */
export function checkChatRow(
	id: string,
	before: unknown,
	after: unknown,
	actor: UpdateActor,
	reverts: RowRevert[],
): UpdateViolation | null {
	const at = { map: CHATS_MAP, key: id };
	if (after !== undefined && !(after instanceof Y.Map)) return { code: 'malformed-row', ...at };
	const next = after as Y.Map<unknown> | undefined;
	if (before === undefined) return next ? checkNewRow(id, next, actor, at) : null;
	if (!next || !(before instanceof Y.Map)) return { code: 'foreign-row', ...at };
	const prior = before as Y.Map<unknown>;
	if (isSameRow(prior, next)) return null;
	if (prior.get('kind') !== 'group') return checkDirectEdit(id, prior, next, actor, at);
	if (memberStanding(prior, actor.id) === 'stranger') return { code: 'not-member', ...at };
	return checkGroupEdit(id, prior, next, actor, at, reverts);
}
