// Identities and comparable content of the entries a peer has written or applied (its baseline).
// An entry is a local edit only when its content differs from the baseline, never merely because
// the room holds something else, so a collaborator's change that is not applied yet is kept.
import { canon } from './package';
import { pieceContent } from './pieces';

/** The scope (the role a page key plays) of the shapes inside master `key`'s content. */
export const masterScope = (key: string): string => `m~${key}`;
export const pageId = (key: string): string => `p:${key}`;
export const bodyId = (key: string): string => `b:${key}`;
export const shapeId = (page: string, key: string): string => `s:${page}/${key}`;
export const cellId = (page: string, shape: string, path: string): string =>
	`f:${page}/${shape}/${path}`;
export const connectId = (key: string): string => `c:${key}`;
export const masterId = (key: string): string => `m:${key}`;
export const mediaId = (key: string): string => `i:${key}`;

/** Baseline text of a page entry (its position excluded) and of a body. */
export const pageContent = (e: { id: string; back?: string | undefined; xml: string }): string =>
	JSON.stringify([e.id, e.back ?? null, canon(e.xml)]);
export const bodyContent = (e: { shell: string; rels?: string | undefined }): string =>
	JSON.stringify([canon(e.shell), e.rels ?? null]);
/** A shape entry: its parent and attributes. */
export const shellContent = (e: { parent?: string | undefined; open: string }): string =>
	JSON.stringify([e.parent ?? null, pieceContent('open', e.open)]);
export const entryContent = (e: { xml?: string | undefined; open?: string | undefined }): string =>
	e.xml !== undefined ? pieceContent('leaf', e.xml) : pieceContent('open', e.open ?? '');
/** A master: id, list element, relationships and a print of its part. */
export const masterContent = (e: {
	id: string;
	xml: string;
	rels?: string | undefined;
	print: string;
}): string => JSON.stringify([e.id, canon(e.xml), e.rels ?? null, e.print]);
