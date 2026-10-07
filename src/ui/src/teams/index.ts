import { defineAppRail } from './app-rail';
import { defineAvatar } from './avatar';
import { defineCallControls } from './call-controls';
import { defineCallGrid } from './call-grid';
import { defineChannelList } from './channel-list';
import { defineChatComposer } from './chat-composer';
import { defineChatList } from './chat-list';
import { definePrejoin } from './prejoin';

export * from './app-rail';
export * from './avatar';
export * from './base';
export * from './call-controls';
export * from './call-grid';
export * from './channel-list';
export * from './chat-composer';
export * from './chat-list';
export * from './icons';
export * from './prejoin';

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

export * from './app/index';
