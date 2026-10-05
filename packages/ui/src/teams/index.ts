import { defineAppRail } from './app-rail.js';
import { defineAvatar } from './avatar.js';
import { defineCallControls } from './call-controls.js';
import { defineCallGrid } from './call-grid.js';
import { defineChannelList } from './channel-list.js';
import { defineChatComposer } from './chat-composer.js';
import { defineChatList } from './chat-list.js';
import { definePrejoin } from './prejoin.js';

export * from './app-rail.js';
export * from './avatar.js';
export * from './base.js';
export * from './call-controls.js';
export * from './call-grid.js';
export * from './channel-list.js';
export * from './chat-composer.js';
export * from './chat-list.js';
export * from './icons.js';
export * from './prejoin.js';

/** Every tag the team-workspace elements define. */
export const TEAMS_TAGS = [
	'office-ui-avatar',
	'office-ui-app-rail',
	'office-ui-channel-list',
	'office-ui-chat-list',
	'office-ui-chat-composer',
	'office-ui-prejoin',
	'office-ui-call-grid',
	'office-ui-call-controls',
] as const;

/** Define the team-workspace elements. Idempotent; a no-op without a DOM. */
export function registerTeams(): void {
	defineAvatar();
	defineAppRail();
	defineChannelList();
	defineChatList();
	defineChatComposer();
	definePrejoin();
	defineCallGrid();
	defineCallControls();
}

export * from './app/index.js';
