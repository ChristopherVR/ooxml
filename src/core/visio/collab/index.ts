// Real-time collaboration for Visio drawings (`ooxml-core/visio/collab`): the Yjs mapping of a
// VSDX package (pages, shapes and connects as keyed entries, every other part verbatim) and the
// binding that keeps an editor and a shared drawing in step. See docs/collab-area.md.
export {
	VISIO_EDIT_ORIGIN,
	VISIO_SYNC_ORIGIN,
	bindDrawingSession,
	type DrawingBinding,
	type DrawingBindingOptions,
	type DrawingCollabHost,
	type DrawingEditor,
} from './bind';
export { composeDrawing, type ComposedDrawing } from './compose';
export { DrawingKeys } from './keys';
export {
	readSharedDrawing,
	type ResolvedDrawing,
	type ResolvedPage,
	type ResolvedShape,
} from './read';
export {
	BODIES,
	CONNECTS,
	DEAD,
	META,
	PAGES,
	PARTS,
	SCHEMA_VERSION,
	SHAPES,
	isSharedDrawingEmpty,
	sharedDrawing,
	type SharedDrawing,
} from './schema';
export { snapshotDrawing, type DrawingSnapshot } from './snapshot';
export { writeSharedDrawing } from './write';
