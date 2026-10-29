// Ownership rules of `validateTeamsUpdate` for the rows that record a creator: channels
// (`teams:channels`), tabs (`teams:tabs`) and pins (`teams:pins:<channel>`). Messages have their
// own rules in validate-update.ts. DOM-free. New code.
import * as Y from 'yjs';
import type { UpdateActor, UpdateViolation } from './validate-update';

const CHANNELS = 'teams:channels';
const TABS = 'teams:tabs';
const MEETINGS = 'teams:meetings';
const CHAT_MEMBERS = 'teams:chat-members';

/** Maps whose rows carry an owner field (`createdBy`, or a meeting's `organizerId`) naming the only non-moderator who may change them. */
export const OWNED_MAPS: readonly string[] = [CHANNELS, TABS, MEETINGS];

/** The row field that names the owner: a meeting's organiser, otherwise the creator. */
const ownerField = (map: string): string => (map === MEETINGS ? 'organizerId' : 'createdBy');

export const isOwnedMap = (name: string): boolean => OWNED_MAPS.includes(name);

/** The same Yjs row with the same content: an insert that lost a concurrent write to its key. */
function isSameRow(before: Y.Map<unknown>, next: Y.Map<unknown>): boolean {
	return (
		next._item !== null &&
		before._item !== null &&
		Y.compareIDs(next._item.id, before._item.id) &&
		JSON.stringify(before.toJSON()) === JSON.stringify(next.toJSON())
	);
}

/**
 * Channel, tab and meeting rows. A new row must name the actor as `createdBy` (a meeting: `organizerId`). Only the creator or a
 * moderator may then rename, archive, replace or remove it, and the creator cannot hand it on.
 * Exception: clients seed a channel with a fixed id (`general`, the `Chat` of a direct chat) when
 * they first connect, so the creator of a channel row may replace another person's row of the same
 * name and kind; changing the name or kind this way is refused.
 */
export function checkOwnedRow(
	map: string,
	rowId: string,
	before: unknown,
	after: unknown,
	actor: UpdateActor,
): UpdateViolation | null {
	const at = { map, key: rowId };
	const field = ownerField(map);
	if (after !== undefined && !(after instanceof Y.Map)) return { code: 'malformed-row', ...at };
	const next = after as Y.Map<unknown> | undefined;
	if (before === undefined)
		return next === undefined || next.get(field) === actor.id
			? null
			: { code: 'forged-author', ...at };
	const prior = before instanceof Y.Map ? before : undefined;
	if (prior && next && isSameRow(prior, next)) return null;
	const owner = prior?.get(field);
	if (owner === actor.id)
		return next === undefined || next.get(field) === owner
			? null
			: { code: 'author-changed', ...at };
	if (actor.moderator === true) return null;
	const replaced = prior !== undefined && next !== undefined;
	if (
		map === CHANNELS &&
		replaced &&
		next.get(field) === actor.id &&
		next.get('name') === prior.get('name') &&
		next.get('kind') === prior.get('kind')
	)
		return null;
	return { code: 'foreign-row', ...at };
}

const pinnedBy = (row: unknown): unknown =>
	row !== null && typeof row === 'object' ? (row as { by?: unknown }).by : undefined;

/**
 * Pin rows. In a private chat (the room has a `teams:chat-members` map with entries) only members
 * may pin or unpin; a new or changed pin must name the actor in `by`. Any member may unpin, as in
 * Teams. A moderator may unpin without being a member.
 */
export function checkPinRow(
	map: string,
	rowId: string,
	before: unknown,
	after: unknown,
	actor: UpdateActor,
	doc: Y.Doc,
): UpdateViolation | null {
	const at = { map, key: rowId };
	if (after !== undefined && (after === null || typeof after !== 'object'))
		return { code: 'malformed-row', ...at };
	const members = doc.share.has(CHAT_MEMBERS) ? doc.getMap<unknown>(CHAT_MEMBERS) : undefined;
	const exempt = actor.moderator === true && after === undefined;
	if (members && members.size > 0 && !members.has(actor.id) && !exempt)
		return { code: 'not-member', ...at };
	if (after === undefined) return null;
	if (pinnedBy(after) === actor.id) return null;
	// Unchanged content (an insert that lost a concurrent write) keeps its original pinner.
	if (before !== undefined && JSON.stringify(before) === JSON.stringify(after)) return null;
	return { code: 'forged-author', ...at };
}
