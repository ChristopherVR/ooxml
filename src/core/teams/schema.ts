// The stored-data contract of a workspace: an integer schema version stamped in the workspace
// document (`teams:meta` map, key `schemaVersion`) and the ordered migrations that bring an older
// document up to it. A client that meets a NEWER schema than it supports does not write: it reports
// `unsupported-schema` and its transport stops sending document changes (see `readOnlyTransport`).
// Compatibility rules: docs/teams-area.md. DOM-free. New code.
import * as decoding from 'lib0/decoding';
import * as Y from 'yjs';
import type { Transport } from '../collab/provider';
import { MESSAGE_SYNC } from '../collab/transport-provider';

/** The schema version this build writes and reads. */
export const SCHEMA_VERSION = 1;

/** Name of the shared map that holds document metadata. */
export const META_MAP = 'teams:meta';
export const SCHEMA_KEY = 'schemaVersion';

/** Transaction origin of migrations and stamps, so a host can tell them from user edits. */
export const MIGRATION_ORIGIN = 'teams:migration';

const CHANNELS = 'teams:channels';
const MESSAGES_PREFIX = 'teams:messages:';

export type SchemaStatus =
	/** Readable and writable. `version` is null while a fresh document has not been stamped yet. */
	| { state: 'ok'; version: number | null; supported: number }
	/** Written by a newer release (`found`) or unreadable (`null`): this client must not write. */
	| { state: 'unsupported-schema'; found: number | null; supported: number };

export interface Migration {
	/** The version this migration produces; it runs on documents at `version - 1`. */
	version: number;
	name: string;
	/** Must be idempotent and only touch what is actually out of shape. Runs in one transaction. */
	up: (doc: Y.Doc) => void;
}

const isPlainMap = (value: unknown): value is Y.Map<unknown> => value instanceof Y.Map;

/**
 * 0 -> 1: documents written before the stamp existed. Give channel and message rows the fields the
 * v1 contract says are always present. Readers keep tolerating a missing field (peers are
 * untrusted); this only makes stored data uniform. A message without `mentions` is valid in v1 and
 * is left alone. The legacy presence `availability` field lives in awareness, not in the document,
 * so it is not a document migration (see `statusFromWire`).
 */
function normaliseV0(doc: Y.Doc): void {
	doc.getMap<unknown>(CHANNELS).forEach((row) => {
		if (!isPlainMap(row)) return;
		if (typeof row.get('topic') !== 'string') row.set('topic', '');
		const kind = row.get('kind');
		if (kind !== 'channel' && kind !== 'direct') row.set('kind', 'channel');
		if (typeof row.get('archived') !== 'boolean') row.set('archived', false);
		if (typeof row.get('createdAt') !== 'number') row.set('createdAt', 0);
	});
	for (const key of [...doc.share.keys()]) {
		if (!key.startsWith(MESSAGES_PREFIX)) continue;
		doc.getMap<unknown>(key).forEach((row) => {
			if (!isPlainMap(row)) return;
			if (typeof row.get('deleted') !== 'boolean') row.set('deleted', false);
			if (typeof row.get('ts') !== 'number') row.set('ts', 0);
		});
	}
}

/** Every migration, in order. Append only: never edit or reorder a shipped one. */
export const MIGRATIONS: readonly Migration[] = [
	{ version: 1, name: 'normalise channel and message rows', up: normaliseV0 },
];

/** The stamped version: a positive integer, null when absent, `NaN` when present but unreadable. */
export function readSchemaVersion(doc: Y.Doc): number | null {
	const raw = doc.getMap<unknown>(META_MAP).get(SCHEMA_KEY);
	if (raw === undefined) return null;
	return typeof raw === 'number' && Number.isInteger(raw) && raw >= 1 ? raw : Number.NaN;
}

/** Whether the document holds any workspace data (a channel), stamped or not. */
export const hasWorkspaceData = (doc: Y.Doc): boolean => doc.getMap(CHANNELS).size > 0;

/** What this build may do with the document. Read only: never writes. */
export function checkSchema(doc: Y.Doc): SchemaStatus {
	const found = readSchemaVersion(doc);
	if (found !== null && Number.isNaN(found))
		return { state: 'unsupported-schema', found: null, supported: SCHEMA_VERSION };
	if (found !== null && found > SCHEMA_VERSION)
		return { state: 'unsupported-schema', found, supported: SCHEMA_VERSION };
	return { state: 'ok', version: found, supported: SCHEMA_VERSION };
}

export interface EnsureSchemaOptions {
	/**
	 * Stamp a document that has no workspace data. Pass false before the first sync, because an
	 * empty local document may just not have received the server's state yet.
	 */
	stampEmpty: boolean;
}

/**
 * Bring the document to {@link SCHEMA_VERSION}: stamp a fresh one, migrate an older one, and leave a
 * newer one untouched. Idempotent, and safe when several clients run it at once (the writes are the
 * same values). Returns the resulting status.
 */
export function ensureSchema(doc: Y.Doc, options: EnsureSchemaOptions): SchemaStatus {
	const status = checkSchema(doc);
	if (status.state !== 'ok' || status.version === SCHEMA_VERSION) return status;
	const from = status.version ?? 0;
	if (status.version === null && !hasWorkspaceData(doc) && !options.stampEmpty) return status;
	doc.transact(() => {
		for (const migration of MIGRATIONS) if (migration.version > from) migration.up(doc);
		doc.getMap<unknown>(META_MAP).set(SCHEMA_KEY, SCHEMA_VERSION);
	}, MIGRATION_ORIGIN);
	return checkSchema(doc);
}

/** Call back with the new status whenever the stamp changes (local or remote). */
export function watchSchema(doc: Y.Doc, onChange: (status: SchemaStatus) => void): () => void {
	const map = doc.getMap<unknown>(META_MAP);
	const listener = (): void => onChange(checkSchema(doc));
	map.observe(listener);
	return () => map.unobserve(listener);
}

/**
 * A transport that sends nothing which would change the server's document while `isReadOnly()`:
 * sync step 2 and update messages are dropped; step 1 (asking for the state) and awareness pass, so
 * the client still reads and shows presence.
 */
export function readOnlyTransport(inner: Transport, isReadOnly: () => boolean): Transport {
	return {
		connect: (handlers) => inner.connect(handlers),
		disconnect: () => inner.disconnect(),
		send(data) {
			if (isReadOnly() && carriesDocumentChanges(data)) return;
			inner.send(data);
		},
	};
}

const SYNC_STEP1 = 0;

function carriesDocumentChanges(data: Uint8Array): boolean {
	try {
		const decoder = decoding.createDecoder(data);
		if (decoding.readVarUint(decoder) !== MESSAGE_SYNC) return false;
		return decoding.readVarUint(decoder) !== SYNC_STEP1;
	} catch {
		return true;
	}
}
