// The shared drawing schema, version 2. Top-level Yjs maps (top-level types never lose data when
// two peers initialise a room at once, unlike nested maps created concurrently). Every entry is a
// plain JSON record replaced atomically per key:
//   visio:meta      v, pagesPath, pagesOpen, pagesClose (the `Pages` part around the page list),
//                   mastersPath, mastersOpen, mastersClose (the same for `Masters`, when present)
//   visio:pages     page key -> PageEntry { pos, id, back?, xml }   (a `Page` element, `Rel` removed)
//   visio:bodies    page key -> BodyEntry { shell, rels?, b? }      (page part without shapes/connects)
//   visio:shapes    `${page key}/${shape key}` -> ShapeEntry { pos, parent?, open, b? }
//                   (one `Shape` opening tag; nested shapes are entries naming their parent)
//   visio:cells     `${page key}/${shape key}/${piece path}` -> PieceEntry { pos, xml | open }
//                   (a Cell, Section, Row or other child of one shape, see pieces.ts)
//   visio:connects  `${page key}/${hash}` -> ConnectEntry                  (one `Connect`)
//   visio:dead      page key | `${page key}/${shape key}` -> true           (tombstones: delete wins)
//   visio:masters   master key -> MasterEntry { pos, id, xml, part, rels?, b? } (one master, whole)
//   visio:media     `${hash}|${path}` -> MediaEntry { path, hash, ct, d, b? }  (one media part)
//   visio:blobs     hash -> base64 bytes of media (identical bytes are stored once)
//   visio:parts     part path -> base64 bytes of every other package part
// Page, shape and master keys are minted by the peer that adds them (`${client}.${n}`) or, when a
// room is seeded, derived from the drawing's own ids (`s${id}`) so two peers seeding the same file
// agree. `b` is the write batch that created or last rewrote an entry: ids a peer wrote in one
// batch share one numbering, which is what lets a renumbering reach the references of its batch.
//
// Version rule: `meta.v` must equal SCHEMA_VERSION for a peer to read or write the room. A peer
// that finds any other version (older or newer) refuses with a SchemaVersionError and stays
// inert for that room: it never writes and never rebuilds the drawing from it. There is no
// migration: version 1 rooms (one entry per whole shape) hold data this schema cannot merge.
import * as Y from 'yjs';
import type { PieceEntry } from './pieces';

export const SCHEMA_VERSION = 2;

/** The room was written by a peer with another schema version. */
export class SchemaVersionError extends Error {
	constructor(readonly found: unknown) {
		super(
			`The shared drawing uses schema version ${String(found)}; this build reads version ${SCHEMA_VERSION}.`,
		);
		this.name = 'SchemaVersionError';
	}
}

export const META = 'visio:meta';
export const PAGES = 'visio:pages';
export const BODIES = 'visio:bodies';
export const SHAPES = 'visio:shapes';
export const CELLS = 'visio:cells';
export const CONNECTS = 'visio:connects';
export const DEAD = 'visio:dead';
export const MASTERS = 'visio:masters';
export const MEDIA = 'visio:media';
export const BLOBS = 'visio:blobs';
export const PARTS = 'visio:parts';

export interface PageEntry {
	pos: number;
	/** The page id (`Page/@ID`) the author gave it; renumbered on read if two pages collide. */
	id: string;
	/** Key of the background page, when the page has one. */
	back?: string;
	/** The `Page` element with its `Rel` child replaced by {@link REL_MARK}. */
	xml: string;
}
export interface BodyEntry {
	/** The page part with `<!--@shapes-->` and `<!--@connects-->` where the lists were. */
	shell: string;
	/** The page part's relationships part, verbatim. */
	rels?: string;
	b?: string;
}
export interface ShapeEntry {
	pos: number;
	/** Key of the shape this one is nested in. */
	parent?: string;
	/** The opening tag of the `Shape` element (its attributes), never self-closing. */
	open: string;
	b?: string;
}
export type CellEntry = PieceEntry;
export interface MasterEntry {
	pos: number;
	/** The master id (`Master/@ID`) the author gave it; renumbered on read if two collide. */
	id: string;
	/** The `Master` element with its `Rel` child replaced by {@link REL_MARK}. */
	xml: string;
	/** The master part, base64. */
	part: string;
	/** The master part's relationships part, verbatim. */
	rels?: string;
	b?: string;
}
export interface MediaEntry {
	path: string;
	hash: string;
	/** Content type declared for the part. */
	ct: string;
	/** Whether that declaration was an extension default rather than an override. */
	d: boolean;
	b?: string;
}
export interface ConnectEntry {
	/** Keys of the connecting and connected shapes ('' keeps the literal id). */
	fk: string;
	fi: string;
	tk: string;
	ti: string;
	fc: string;
	tc: string;
	fp: string;
	tp: string;
}

export const REL_MARK = '<!--@rel-->';
export const SHAPES_MARK = '<!--@shapes-->';
export const CONNECTS_MARK = '<!--@connects-->';

export interface SharedDrawing {
	meta: Y.Map<unknown>;
	pages: Y.Map<unknown>;
	bodies: Y.Map<unknown>;
	shapes: Y.Map<unknown>;
	cells: Y.Map<unknown>;
	connects: Y.Map<unknown>;
	dead: Y.Map<unknown>;
	masters: Y.Map<unknown>;
	media: Y.Map<unknown>;
	blobs: Y.Map<unknown>;
	parts: Y.Map<unknown>;
}

/** The drawing's top-level maps in `doc`. */
export function sharedDrawing(doc: Y.Doc): SharedDrawing {
	return {
		meta: doc.getMap<unknown>(META),
		pages: doc.getMap<unknown>(PAGES),
		bodies: doc.getMap<unknown>(BODIES),
		shapes: doc.getMap<unknown>(SHAPES),
		cells: doc.getMap<unknown>(CELLS),
		connects: doc.getMap<unknown>(CONNECTS),
		dead: doc.getMap<unknown>(DEAD),
		masters: doc.getMap<unknown>(MASTERS),
		media: doc.getMap<unknown>(MEDIA),
		blobs: doc.getMap<unknown>(BLOBS),
		parts: doc.getMap<unknown>(PARTS),
	};
}

/** Every top-level map, for an UndoManager scope. */
export const sharedTypes = (shared: SharedDrawing): Y.Map<unknown>[] => [
	shared.meta,
	shared.pages,
	shared.bodies,
	shared.shapes,
	shared.cells,
	shared.connects,
	shared.dead,
	shared.masters,
	shared.media,
	shared.blobs,
	shared.parts,
];

/** True when no peer has written a drawing to the room yet. */
export const isSharedDrawingEmpty = (shared: SharedDrawing): boolean =>
	shared.pages.size === 0 && shared.meta.get('v') === undefined;

/** Throws a {@link SchemaVersionError} when the room holds another schema version. */
export function assertSchemaVersion(shared: SharedDrawing): void {
	const v = shared.meta.get('v');
	if (v !== undefined && v !== SCHEMA_VERSION) throw new SchemaVersionError(v);
}

export const isRecord = (value: unknown): value is Record<string, unknown> =>
	typeof value === 'object' && value !== null && !Array.isArray(value);

/** Whether two flat records hold the same fields. */
export function sameRecord(a: unknown, b: unknown): boolean {
	if (!isRecord(a) || !isRecord(b)) return a === b;
	const ka = Object.keys(a).filter((k) => a[k] !== undefined);
	const kb = Object.keys(b).filter((k) => b[k] !== undefined);
	return ka.length === kb.length && ka.every((k) => a[k] === b[k]);
}

/** Sets `key` only when its value differs; returns whether it wrote. */
export function setIfChanged(map: Y.Map<unknown>, key: string, value: unknown): boolean {
	if (sameRecord(map.get(key), value)) return false;
	map.set(key, value);
	return true;
}

export const str = (value: unknown): string | undefined =>
	typeof value === 'string' ? value : undefined;
export const num = (value: unknown): number | undefined =>
	typeof value === 'number' && Number.isFinite(value) ? value : undefined;
const opt = <K extends string>(key: K, value: string | undefined): { [P in K]?: string } =>
	value === undefined ? {} : ({ [key]: value } as { [P in K]?: string });

/** Narrowers for entries read from the room (a peer may have written anything). */
export function pageEntry(value: unknown): PageEntry | undefined {
	if (!isRecord(value)) return undefined;
	const pos = num(value.pos);
	const id = str(value.id);
	const xml = str(value.xml);
	if (pos === undefined || id === undefined || xml === undefined) return undefined;
	return { pos, id, xml, ...opt('back', str(value.back)) };
}
export function bodyEntry(value: unknown): BodyEntry | undefined {
	if (!isRecord(value)) return undefined;
	const shell = str(value.shell);
	if (shell === undefined) return undefined;
	return { shell, ...opt('rels', str(value.rels)), ...opt('b', str(value.b)) };
}
export function shapeEntry(value: unknown): ShapeEntry | undefined {
	if (!isRecord(value)) return undefined;
	const pos = num(value.pos);
	const open = str(value.open);
	if (pos === undefined || open === undefined) return undefined;
	return { pos, open, ...opt('parent', str(value.parent)), ...opt('b', str(value.b)) };
}
export function cellEntry(value: unknown): CellEntry | undefined {
	if (!isRecord(value)) return undefined;
	const pos = num(value.pos);
	const xml = str(value.xml);
	const open = str(value.open);
	if (pos === undefined || (xml === undefined && open === undefined)) return undefined;
	return { pos, ...opt('xml', xml), ...opt('open', xml === undefined ? open : undefined) };
}
export function masterEntry(value: unknown): MasterEntry | undefined {
	if (!isRecord(value)) return undefined;
	const pos = num(value.pos);
	const id = str(value.id);
	const xml = str(value.xml);
	const part = str(value.part);
	if (pos === undefined || id === undefined || xml === undefined || part === undefined)
		return undefined;
	return { pos, id, xml, part, ...opt('rels', str(value.rels)), ...opt('b', str(value.b)) };
}
export function mediaEntry(value: unknown): MediaEntry | undefined {
	if (!isRecord(value)) return undefined;
	const path = str(value.path);
	const hash = str(value.hash);
	const ct = str(value.ct);
	if (path === undefined || hash === undefined || ct === undefined) return undefined;
	return { path, hash, ct, d: value.d === true, ...opt('b', str(value.b)) };
}
export function connectEntry(value: unknown): ConnectEntry | undefined {
	if (!isRecord(value)) return undefined;
	const out: Record<string, string> = {};
	for (const field of ['fk', 'fi', 'tk', 'ti', 'fc', 'tc', 'fp', 'tp']) {
		const v = str(value[field]);
		if (v === undefined) return undefined;
		out[field] = v;
	}
	return out as unknown as ConnectEntry;
}
