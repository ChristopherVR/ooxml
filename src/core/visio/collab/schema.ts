// The shared drawing schema. Seven top-level Yjs maps (top-level types never lose data when two
// peers initialise a room at once, unlike nested maps created concurrently). Every entry is a plain
// JSON record replaced atomically per key:
//   visio:meta      v, pagesPath, pagesOpen, pagesClose (the `Pages` part around the page list)
//   visio:pages     page key -> PageEntry { pos, id, back?, xml }   (a `Page` element, `Rel` removed)
//   visio:bodies    page key -> BodyEntry { shell, rels? }          (page part without shapes/connects)
//   visio:shapes    `${page key}/${shape key}` -> ShapeEntry { pos, xml }   (one top-level `Shape`)
//   visio:connects  `${page key}/${hash}` -> ConnectEntry                  (one `Connect`)
//   visio:dead      page key | `${page key}/${shape key}` -> true           (tombstones: delete wins)
//   visio:parts     part path -> base64 bytes of every other package part
// Page and shape keys are minted by the peer that adds them (`${client}.${n}`) or, when a room is
// seeded, derived from the drawing's own ids (`s${id}`) so two peers seeding the same file agree.
import * as Y from 'yjs';

export const SCHEMA_VERSION = 1;

export const META = 'visio:meta';
export const PAGES = 'visio:pages';
export const BODIES = 'visio:bodies';
export const SHAPES = 'visio:shapes';
export const CONNECTS = 'visio:connects';
export const DEAD = 'visio:dead';
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
}
export interface ShapeEntry {
	pos: number;
	xml: string;
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
	connects: Y.Map<unknown>;
	dead: Y.Map<unknown>;
	parts: Y.Map<unknown>;
}

/** The drawing's top-level maps in `doc`. */
export function sharedDrawing(doc: Y.Doc): SharedDrawing {
	return {
		meta: doc.getMap<unknown>(META),
		pages: doc.getMap<unknown>(PAGES),
		bodies: doc.getMap<unknown>(BODIES),
		shapes: doc.getMap<unknown>(SHAPES),
		connects: doc.getMap<unknown>(CONNECTS),
		dead: doc.getMap<unknown>(DEAD),
		parts: doc.getMap<unknown>(PARTS),
	};
}

/** Every top-level map, for an UndoManager scope. */
export const sharedTypes = (shared: SharedDrawing): Y.Map<unknown>[] => [
	shared.meta,
	shared.pages,
	shared.bodies,
	shared.shapes,
	shared.connects,
	shared.dead,
	shared.parts,
];

/** True when no peer has written a drawing to the room yet. */
export const isSharedDrawingEmpty = (shared: SharedDrawing): boolean =>
	shared.pages.size === 0 && shared.meta.get('v') === undefined;

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

/** Narrowers for entries read from the room (a peer may have written anything). */
export function pageEntry(value: unknown): PageEntry | undefined {
	if (!isRecord(value)) return undefined;
	const pos = num(value.pos);
	const id = str(value.id);
	const xml = str(value.xml);
	if (pos === undefined || id === undefined || xml === undefined) return undefined;
	const back = str(value.back);
	return { pos, id, xml, ...(back !== undefined ? { back } : {}) };
}
export function bodyEntry(value: unknown): BodyEntry | undefined {
	if (!isRecord(value)) return undefined;
	const shell = str(value.shell);
	if (shell === undefined) return undefined;
	const rels = str(value.rels);
	return { shell, ...(rels !== undefined ? { rels } : {}) };
}
export function shapeEntry(value: unknown): ShapeEntry | undefined {
	if (!isRecord(value)) return undefined;
	const pos = num(value.pos);
	const xml = str(value.xml);
	return pos === undefined || xml === undefined ? undefined : { pos, xml };
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
