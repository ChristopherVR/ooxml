// `ooxml-core/teams`: the logic of an open-source, bring-your-own-server team workspace that sits
// beside Word, Excel, PowerPoint and Visio: channels and chat on Yjs, presence, WebRTC calls with
// pluggable signaling, and server configuration. No DOM, no UI framework. See docs/teams-area.md.
export * from './call.js';
export * from './content.js';
export type { FileOperationOptions, FileTransferProgress } from './file-transfer.js';
export * from './chat.js';
export * from './model.js';
export * from './peer.js';
export * from './server-config.js';
export * from './signaling.js';
export * from './store.js';
export * from './tabs.js';
export * from './threads.js';
export type { FollowedThread, ThreadFollowSettings } from './followed-threads.js';
export * from './view.js';
export * from './workspace.js';
