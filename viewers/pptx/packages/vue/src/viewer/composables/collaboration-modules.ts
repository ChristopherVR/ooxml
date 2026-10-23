import type { CollabYjsModules } from 'ooxml-ui/pptx';

/**
 * This package's lazy loaders for the optional `yjs` / `y-websocket` /
 * `y-webrtc` peers, handed to the shared provider and session factories so a
 * host that never collaborates never loads them.
 */
export const collabModules: CollabYjsModules = {
	loadYjs: () => import('yjs'),
	loadWebsocket: () => import('y-websocket'),
	loadWebrtc: () => import('y-webrtc'),
};
