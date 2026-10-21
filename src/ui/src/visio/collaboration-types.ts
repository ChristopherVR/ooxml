// The collaboration contract of `<visio-viewer>`: the shared embedding contract of
// `ooxml-core/collab`, named for this product. Visio has no presence payload of its own, so a
// host-made session is a plain `CollabSession`.
import type { CollaborationOptions, CollaborationState } from 'ooxml-core/collab';

export type VisioCollaborationOptions = CollaborationOptions;
export type VisioCollaborationState = CollaborationState;
