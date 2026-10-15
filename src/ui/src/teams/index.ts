import { defineActivityFeed } from './activity-feed';
import { defineAppRail } from './app-rail';
import { defineAvatar } from './avatar';
import { defineCallControls } from './call-controls';
import { defineCalendarAgenda } from './calendar-agenda';
import { defineCalendarView } from './calendar-view';
import { defineCallGrid } from './call-grid';
import { defineChannelList } from './channel-list';
import { defineChatComposer } from './chat-composer';
import { defineChatList } from './chat-list';
import { defineChatSidebar } from './chat-sidebar';
import { definePeoplePicker } from './people-picker';
import { definePrejoin } from './prejoin';

export * from './activity-feed';
export * from './app-rail';
export * from './avatar';
export * from './base';
export * from './call-controls';
export * from './calendar-agenda';
export * from './calendar-view';
export * from './call-grid';
export * from './channel-list';
export * from './chat-composer';
export * from './chat-list';
export * from './chat-sidebar';
export * from './people-picker';
export * from './icons';
export * from './prejoin';

/** Every tag the team-workspace elements define. */
export const TEAMS_TAGS = [
	'office-ui-avatar',
	'office-ui-app-rail',
	'office-ui-activity-feed',
	'office-ui-channel-list',
	'office-ui-chat-list',
	'office-ui-chat-composer',
	'office-ui-prejoin',
	'office-ui-call-grid',
	'office-ui-call-controls',
	'office-ui-chat-sidebar',
	'office-ui-people-picker',
	'office-ui-calendar-view',
	'office-ui-calendar-agenda',
] as const;

/** Define the team-workspace elements. Idempotent; a no-op without a DOM. */
export function registerTeams(): void {
	defineAvatar();
	defineAppRail();
	defineActivityFeed();
	defineChannelList();
	defineChatList();
	defineChatComposer();
	definePrejoin();
	defineCallGrid();
	defineCallControls();
	defineChatSidebar();
	definePeoplePicker();
	defineCalendarView();
	defineCalendarAgenda();
}

export * from './app/index';
