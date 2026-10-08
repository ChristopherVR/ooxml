// `ooxml-core/teams`: the logic of an open-source, bring-your-own-server team workspace that sits
// beside Word, Excel, PowerPoint and Visio: channels and chat on Yjs, presence, WebRTC calls with
// pluggable signaling, and server configuration. No DOM, no UI framework. See docs/teams-area.md.
export * from './call';
export * from './content';
export type { FileOperationOptions, FileTransferProgress } from './file-transfer';
export * from './chat';
export * from './model';
export * from './peer';
export * from './server-config';
export * from './server-file-storage';
export * from './signaling';
export * from './store';
export * from './tabs';
export * from './tab-conversation';
export * from './threads';
export type { ChatDraft, DraftContext, SavedDraft } from './drafts';
export type { MessageTransfer } from './message-transfer';
export type { FollowedThread, ThreadFollowSettings } from './followed-threads';
export * from './view';
export * from './file-order';
export * from './workspace';
