// Server-side authorship checks for Yjs updates. A client alone cannot enforce "you may only post
// as yourself and delete your own messages", because a hostile client writes whatever the server
// accepts. `validateTeamsUpdate` answers, before the server applies an update to a room document:
// would this update, applied on top of the current state, make the authenticated actor forge or
// destroy somebody else's message or reaction? The room document is never modified. What this does
// and does not cover is listed in docs/teams-area.md ("Server-side authorship"). DOM-free. New code.
import * as decoding from 'lib0/decoding';
import * as Y from 'yjs';
import { MAX_UPDATE_BYTES } from '../collab/codec';
import { MESSAGE_SYNC } from '../collab/transport-provider';

const MESSAGES_PREFIX = 'teams:messages:';
const REACTIONS = 'teams:reactions';
const REACTION_SEP = '|';
/** Tab discussion roots have deterministic ids; concurrent clients legitimately write the same one. */
const TAB_ROOT_PREFIX = 'tab:';

export interface UpdateActor {
	/** The authenticated user id (from a signed session), never a claimed one. */
	id: string;
	/** May hide (soft-delete or remove) anybody's message and remove anybody's reaction. */
	moderator?: boolean;
}

export type UpdateViolationCode =
	/** A new message row names an author other than the actor. */
	| 'forged-author'
	/** A message row of somebody else was edited, replaced or removed. */
	| 'foreign-row'
	/** The author of an existing row was changed. */
	| 'author-changed'
	/** A message row is not a map, or a reaction key has the wrong shape. */
	| 'malformed-row'
	/** A reaction was added or removed in somebody else's name. */
	| 'foreign-reaction'
	/** The update refers to state the server does not have, so it cannot be checked. */
	| 'incomplete-update'
	/** The bytes are not a valid Yjs update, or are too large. */
	| 'invalid-update';

export interface UpdateViolation {
	code: UpdateViolationCode;
	/** The message map (`teams:messages:<channel>`) or `teams:reactions`. */
	map?: string;
	/** The message id or reaction key. */
	key?: string;
}

export type UpdateVerdict = { ok: true } | { ok: false; violations: UpdateViolation[] };

const OK: UpdateVerdict = { ok: true };
const fail = (...violations: UpdateViolation[]): UpdateVerdict => ({ ok: false, violations });

/** The map a document holds under `name`, or undefined (the document is not given new roots). */
function mapOf(doc: Y.Doc, name: string): Y.Map<unknown> | undefined {
	if (!doc.share.has(name)) return undefined;
	try {
		return doc.getMap<unknown>(name);
	} catch {
		return undefined;
	}
}

/** What the walk needs of a shared type; its event type parameter is irrelevant here. */
interface TypeNode {
	_item: Y.Item | null;
}

function rootName(doc: Y.Doc, type: TypeNode): string | undefined {
	for (const [name, root] of doc.share) if ((root as TypeNode) === type) return name;
	return undefined;
}

interface Touched {
	rows: Map<string, Set<string>>;
	reactions: Set<string>;
}

/** Which message rows and reaction keys a transaction changed, including edits nested in a row. */
function collectTouched(doc: Y.Doc, txn: Y.Transaction, into: Touched): void {
	for (const [type, keys] of txn.changed) {
		if (type._item === null) {
			const name = rootName(doc, type);
			if (name === undefined) continue;
			for (const key of keys) {
				if (key === null) continue;
				if (name.startsWith(MESSAGES_PREFIX)) {
					let rows = into.rows.get(name);
					if (!rows) into.rows.set(name, (rows = new Set()));
					rows.add(key);
				} else if (name === REACTIONS) into.reactions.add(key);
			}
			continue;
		}
		// A change inside a row (or deeper): climb to the row that sits in a message map.
		let node: TypeNode = type;
		while (node._item && node._item.parent instanceof Y.AbstractType) {
			const parent: TypeNode = node._item.parent;
			if (parent._item === null) {
				const name = rootName(doc, parent);
				const rowId = node._item.parentSub;
				if (name?.startsWith(MESSAGES_PREFIX) && rowId !== null) {
					let rows = into.rows.get(name);
					if (!rows) into.rows.set(name, (rows = new Set()));
					rows.add(rowId);
				}
				break;
			}
			node = parent;
		}
	}
}

const authorOf = (row: unknown): unknown =>
	row instanceof Y.Map ? row.get('authorId') : undefined;

/** Only `deleted` differs, and it became true: what a moderator may do to somebody's row. */
function isHideOnly(before: Y.Map<unknown>, after: Y.Map<unknown>): boolean {
	if (after.get('deleted') !== true) return false;
	const keys = new Set([...before.keys(), ...after.keys()]);
	for (const key of keys) {
		if (key === 'deleted') continue;
		if (JSON.stringify(before.get(key)) !== JSON.stringify(after.get(key))) return false;
	}
	return true;
}

function checkRow(
	map: string,
	rowId: string,
	before: unknown,
	after: unknown,
	actor: UpdateActor,
): UpdateViolation | null {
	const at = { map, key: rowId };
	if (after !== undefined && !(after instanceof Y.Map)) return { code: 'malformed-row', ...at };
	const next = after as Y.Map<unknown> | undefined;
	if (before === undefined) {
		// Created (or created and removed again in the same update).
		return next === undefined || next.get('authorId') === actor.id
			? null
			: { code: 'forged-author', ...at };
	}
	const owner = authorOf(before);
	// The same row with the same content: an insert that lost a concurrent write to the same key.
	if (
		next !== undefined &&
		before instanceof Y.Map &&
		next._item &&
		before._item &&
		Y.compareIDs(next._item.id, before._item.id) &&
		JSON.stringify(before.toJSON()) === JSON.stringify(next.toJSON())
	)
		return null;
	if (owner === actor.id) {
		// Your own row: edit, soft delete or remove it, but do not hand it to somebody else.
		return next === undefined || next.get('authorId') === owner
			? null
			: { code: 'author-changed', ...at };
	}
	const replaced =
		next !== undefined &&
		before instanceof Y.Map &&
		!(next._item && before._item && Y.compareIDs(next._item.id, before._item.id));
	// Concurrent starts of the same tab discussion write one deterministic id; any signed-in user
	// may (re)write it as themselves. Everything else of somebody else's is off limits.
	if (replaced && rowId.startsWith(TAB_ROOT_PREFIX) && next?.get('authorId') === actor.id)
		return null;
	if (actor.moderator === true) {
		if (next === undefined) return null;
		if (!replaced && before instanceof Y.Map && isHideOnly(before, next)) return null;
	}
	return { code: 'foreign-row', ...at };
}

function checkReaction(
	key: string,
	beforeHas: boolean,
	afterHas: boolean,
	actor: UpdateActor,
): UpdateViolation | null {
	const parts = key.split(REACTION_SEP);
	if (parts.length !== 3 || !parts[2]) return { code: 'malformed-row', map: REACTIONS, key };
	if (parts[2] === actor.id) return null;
	// A moderator may remove somebody's reaction, never add one in their name.
	if (actor.moderator === true && beforeHas && !afterHas) return null;
	return { code: 'foreign-reaction', map: REACTIONS, key };
}

/**
 * Check `update` against the current state of `doc` without changing `doc`. The update is applied
 * to a scratch copy; every message row and reaction it touches is compared with the original.
 * Rules, for a non-moderator actor:
 * - a new message row must have `authorId` equal to the actor;
 * - an existing row may be edited, soft-deleted or removed only by its author, who cannot change
 *   `authorId`;
 * - a reaction key `message|emoji|user` may be added or removed only with `user` equal to the actor.
 * A moderator may in addition set `deleted` to true on a row or remove a row, and remove a reaction.
 * Everything else in the document (channels, pins, tabs, meetings, roster, chat members) passes;
 * see the coverage list in docs/teams-area.md.
 *
 * An update that depends on state the server lacks stays pending in Yjs and would be applied later
 * unchecked, so it is refused as `incomplete-update`.
 */
export function validateTeamsUpdate(
	doc: Y.Doc,
	update: Uint8Array,
	actor: UpdateActor,
	options: { maxBytes?: number } = {},
): UpdateVerdict {
	if (!(update instanceof Uint8Array) || update.byteLength > (options.maxBytes ?? MAX_UPDATE_BYTES))
		return fail({ code: 'invalid-update' });
	const scratch = new Y.Doc();
	const touched: Touched = { rows: new Map(), reactions: new Set() };
	try {
		Y.applyUpdate(scratch, Y.encodeStateAsUpdate(doc));
		scratch.on('afterTransaction', (txn: Y.Transaction) => collectTouched(scratch, txn, touched));
		Y.applyUpdate(scratch, update);
	} catch {
		scratch.destroy();
		return fail({ code: 'invalid-update' });
	}
	try {
		if (scratch.store.pendingStructs !== null || scratch.store.pendingDs !== null)
			return fail({ code: 'incomplete-update' });
		const violations: UpdateViolation[] = [];
		for (const [name, rows] of touched.rows) {
			const beforeMap = mapOf(doc, name);
			const afterMap = mapOf(scratch, name);
			for (const rowId of rows) {
				const found = checkRow(name, rowId, beforeMap?.get(rowId), afterMap?.get(rowId), actor);
				if (found) violations.push(found);
			}
		}
		const beforeReactions = mapOf(doc, REACTIONS);
		const afterReactions = mapOf(scratch, REACTIONS);
		for (const key of touched.reactions) {
			const found = checkReaction(
				key,
				beforeReactions?.has(key) === true,
				afterReactions?.has(key) === true,
				actor,
			);
			if (found) violations.push(found);
		}
		return violations.length ? { ok: false, violations } : OK;
	} finally {
		scratch.destroy();
	}
}

/**
 * The same check for a raw y-websocket message as the server receives it: sync step 2 and update
 * messages are validated; step 1, awareness and everything else pass. Use it right before handing
 * the message to `readSyncMessage`.
 */
export function validateTeamsSyncMessage(
	doc: Y.Doc,
	message: Uint8Array,
	actor: UpdateActor,
): UpdateVerdict {
	let payload: Uint8Array;
	try {
		const decoder = decoding.createDecoder(message);
		if (decoding.readVarUint(decoder) !== MESSAGE_SYNC) return OK;
		const kind = decoding.readVarUint(decoder);
		if (kind !== 1 && kind !== 2) return OK;
		payload = decoding.readVarUint8Array(decoder);
	} catch {
		return fail({ code: 'invalid-update' });
	}
	return validateTeamsUpdate(doc, payload, actor);
}
