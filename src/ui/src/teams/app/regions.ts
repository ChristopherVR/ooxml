/**
 * The workspace's F6 ring: the app bar, the channel or chat list, the channel header, the message
 * list and the composer, in the order Microsoft Teams moves round the window. The walk is the
 * shared region service every Office product uses, so F6 behaves the same here as in the editors.
 */
import {
	moveRegionFocus,
	regionKeyDirection,
	type RegionScope,
	type WindowRegion,
} from '../../region-focus';

const TEAMS_REGIONS: readonly WindowRegion[] = [
	{ selector: 'office-ui-app-rail' },
	{ selector: 'office-ui-channel-list, office-ui-chat-sidebar' },
	{ selector: 'header.channel-head' },
	{ selector: 'office-ui-chat-list' },
	// The composer hands the keyboard to its message box, not to Attach or Send.
	{ selector: 'office-ui-chat-composer', entry: 'textarea' },
];

/**
 * Answers F6 and Shift+F6 on the app shell: moves the keyboard to the next or previous region and
 * says whether it did, so the caller can keep the key from the browser.
 */
export function moveTeamsRegion(shell: RegionScope, event: KeyboardEvent): boolean {
	const direction = regionKeyDirection(event);
	return direction !== undefined && !event.defaultPrevented
		? moveRegionFocus(shell, TEAMS_REGIONS, direction)
		: false;
}
