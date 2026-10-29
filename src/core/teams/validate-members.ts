// Rules of `validateTeamsUpdate` for the roster (`teams:roster`, one row per user id) and the chat
// member map (`teams:chat-members`, user id to true, the list the server admits sockets by), plus
// the content limits of message rows and reactions. DOM-free. New code.
import * as Y from 'yjs';
import type { UpdateActor, UpdateViolation } from './validate-update';

export const ROSTER = 'teams:roster';
export const CHAT_MEMBERS = 'teams:chat-members';

/** Longest message text (UTF-16 units) a row may carry; the client caps far lower. */
export const DEFAULT_MAX_MESSAGE_CHARS = 32_768;
/** Most reaction keys one message may collect (emoji times people). */
export const DEFAULT_MAX_REACTIONS_PER_MESSAGE = 200;

export interface ContentLimits {
	maxMessageChars?: number;
	maxReactionsPerMessage?: number;
}

/**
 * A roster row is the user's own entry: its key must be the actor's id, for writing, refreshing
 * and removing. A moderator may remove anybody's entry (never write one in their name).
 */
export function checkRosterRow(
	rowId: string,
	before: unknown,
	after: unknown,
	actor: UpdateActor,
): UpdateViolation | null {
	const at = { map: ROSTER, key: rowId };
	if (after !== undefined && !(after instanceof Y.Map)) return { code: 'malformed-row', ...at };
	if (rowId === actor.id) return null;
	if (actor.moderator === true && after === undefined && before !== undefined) return null;
	return { code: 'foreign-row', ...at };
}

/**
 * The member map of a private chat. Adding a member needs the actor to be a member already (the
 * first writer of an empty map must list themselves). A member may remove only themselves; a
 * moderator may remove anyone.
 */
export function checkMemberKey(
	userId: string,
	before: Y.Map<unknown> | undefined,
	after: Y.Map<unknown> | undefined,
	actor: UpdateActor,
): UpdateViolation | null {
	const at = { map: CHAT_MEMBERS, key: userId };
	const was = before?.has(userId) === true;
	const now = after?.has(userId) === true;
	if (was === now) return null;
	if (now) {
		if (after?.get(userId) !== true) return { code: 'malformed-row', ...at };
		if (actor.moderator === true) return null;
		const existing = before !== undefined && before.size > 0;
		const allowed = existing ? before.has(actor.id) : after?.has(actor.id) === true;
		return allowed ? null : { code: 'not-member', ...at };
	}
	return userId === actor.id || actor.moderator === true ? null : { code: 'foreign-row', ...at };
}

/** A message row's text must stay within the limit. */
export function checkMessageSize(
	map: string,
	rowId: string,
	after: unknown,
	limits: ContentLimits,
): UpdateViolation | null {
	if (!(after instanceof Y.Map)) return null;
	const text = after.get('text');
	const max = limits.maxMessageChars ?? DEFAULT_MAX_MESSAGE_CHARS;
	return typeof text === 'string' && text.length > max
		? { code: 'content-too-large', map, key: rowId }
		: null;
}

/** The reactions on one message (keys `message|emoji|user`) must stay within the limit. */
export function checkReactionCount(
	messageId: string,
	reactions: Y.Map<unknown> | undefined,
	limits: ContentLimits,
): UpdateViolation | null {
	if (!reactions) return null;
	const prefix = `${messageId}|`;
	let count = 0;
	for (const key of reactions.keys()) if (key.startsWith(prefix)) count++;
	return count > (limits.maxReactionsPerMessage ?? DEFAULT_MAX_REACTIONS_PER_MESSAGE)
		? { code: 'content-too-large', map: 'teams:reactions', key: messageId }
		: null;
}
