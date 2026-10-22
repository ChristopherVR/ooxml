// The collaboration contract of `<visio-viewer>`: the shared embedding contract of
// `ooxml-core/collab`, named for this product, with the Visio presence payload and the owner's
// write-back hook (defined in `ooxml-core/visio/ui`, where the bindings take it from).
import type { CollaborationState } from 'ooxml-core/collab';

export type {
	VisioCollaborationOptions,
	VisioPresence,
	VisioWriteBackOptions,
} from 'ooxml-core/visio/ui';
export type VisioCollaborationState = CollaborationState;
