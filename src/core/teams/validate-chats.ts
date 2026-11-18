// The rule of `validateTeamsUpdate` for the chat directory (`teams:chats`, one row per chat in the
// workspace document). A row says who is in a chat, so only that chat's members may change it, and
// what kind of chat it is and who created it never change. DOM-free.
import * as Y from 'yjs';
import {
	CHATS_MAP,
	type ChatInfo,
	directChatId,
	isGroupChatIdOf,
	readChatRow,
} from './chat-records';
import type { UpdateActor, UpdateViolation } from './validate-update';
import { isSameRow } from './validate-owned';

const IMMUTABLE_FIELDS = ['kind', 'createdAt', 'createdBy', 'meetingId'];

/** The row as a chat, or null when a reader would have to drop any part of it. */
function strictChat(id: string, row: Y.Map<unknown>): ChatInfo | null {
	const chat = readChatRow(id, row);
	const raw = row.get('members');
	return chat && Array.isArray(raw) && raw.length === chat.members.length ? chat : null;
}

const memberIds = (chat: ChatInfo): Set<string> => new Set(chat.members.map((m) => m.id));
const sameIds = (a: Set<string>, b: Set<string>): boolean =>
	a.size === b.size && [...a].every((id) => b.has(id));

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

/**
 * One directory row. A new row names the actor as creator and member, under an id bound to them: a
 * group id they minted, or the 1:1 id of the pair. After that `kind`, `createdAt`, `createdBy` and
 * `meetingId` never change and a row is never removed. A 1:1 row is immutable, except that either
 * of its two people may replace it with an equal one, because both can open the chat at once. In a
 * group only a current member may rename or add people, and a member removes only themselves (a
 * moderator who is a member removes anyone).
 */
export function checkChatRow(
	id: string,
	before: unknown,
	after: unknown,
	actor: UpdateActor,
): UpdateViolation | null {
	const at = { map: CHATS_MAP, key: id };
	if (after !== undefined && !(after instanceof Y.Map)) return { code: 'malformed-row', ...at };
	const next = after as Y.Map<unknown> | undefined;
	if (before === undefined) return next ? checkNewRow(id, next, actor, at) : null;
	if (!next) return { code: 'foreign-row', ...at };
	const prior = before instanceof Y.Map ? before : undefined;
	if (prior && isSameRow(prior, next)) return null;
	const was = prior ? readChatRow(id, prior) : null;
	if (!prior || !was) return { code: 'foreign-row', ...at };
	const had = memberIds(was);
	if (!had.has(actor.id)) return { code: 'not-member', ...at };
	const now = strictChat(id, next);
	if (!now || now.kind !== was.kind) return { code: 'malformed-row', ...at };
	const has = memberIds(now);
	if (was.kind === 'direct') {
		const replaced = !(next._item && prior._item && Y.compareIDs(next._item.id, prior._item.id));
		return replaced && next.get('createdBy') === actor.id && sameIds(had, has)
			? null
			: { code: 'foreign-row', ...at };
	}
	if (IMMUTABLE_FIELDS.some((field) => next.get(field) !== prior.get(field)))
		return { code: 'foreign-row', ...at };
	for (const member of had)
		if (!has.has(member) && member !== actor.id && actor.moderator !== true)
			return { code: 'foreign-row', ...at };
	return null;
}
