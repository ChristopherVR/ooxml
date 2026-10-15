import type { PinnedMessage } from 'ooxml-core/teams';
import { TeamsPinnedBar, defineTeamsPinnedBar } from './pinned-bar';
import { OfficeUiChatList, defineChatList } from '../chat-list';
import { defineAvatar } from '../avatar';

const pin = (id: string, author: string, text: string, by = 'Bob'): PinnedMessage => ({
	message: {
		id,
		channelId: 'general',
		authorId: author.toLowerCase(),
		authorName: author,
		text,
		ts: 1,
		deleted: false,
		attachments: [],
		reactions: {},
	},
	pinnedBy: by.toLowerCase(),
	pinnedByName: by,
	pinnedAt: 1,
});

beforeAll(() => {
	defineTeamsPinnedBar();
	defineAvatar();
	defineChatList();
});
afterEach(() => document.body.replaceChildren());

describe('<teams-pinned-bar>', () => {
	it('renders nothing without pins and the newest pin with a count otherwise', async () => {
		const bar = document.createElement('teams-pinned-bar') as TeamsPinnedBar;
		document.body.append(bar);
		await bar.updateComplete;
		expect(bar.shadowRoot!.querySelector('section')).toBeNull();
		bar.pinned = [pin('m1', 'Ada', 'Ship it'), pin('m2', 'Bob', 'Old note')];
		await bar.updateComplete;
		expect(bar.shadowRoot!.querySelector('.latest')!.textContent).toContain('Ship it');
		expect(bar.shadowRoot!.querySelector('.count')!.textContent!.trim()).toBe('2 pinned');
		expect(bar.shadowRoot!.querySelector('ul')).toBeNull();
	});

	it('expands the list and emits open and unpin events', async () => {
		const bar = document.createElement('teams-pinned-bar') as TeamsPinnedBar;
		bar.pinned = [pin('m1', 'Ada', 'Ship it'), pin('m2', 'Bob', 'Old note')];
		document.body.append(bar);
		await bar.updateComplete;
		const seen: string[] = [];
		bar.addEventListener('teams-pinned-open', (e) =>
			seen.push(`open:${(e as CustomEvent).detail.messageId}`),
		);
		bar.addEventListener('teams-pinned-unpin', (e) =>
			seen.push(`unpin:${(e as CustomEvent).detail.messageId}`),
		);
		bar.shadowRoot!.querySelector<HTMLButtonElement>('.count')!.click();
		await bar.updateComplete;
		expect(bar.shadowRoot!.querySelectorAll('li')).toHaveLength(2);
		expect(bar.shadowRoot!.querySelector('.by')!.textContent).toBe('Pinned by Bob');
		bar.shadowRoot!.querySelectorAll<HTMLButtonElement>('.jump')[1]!.click();
		bar.shadowRoot!.querySelectorAll<HTMLButtonElement>('.unpin')[0]!.click();
		expect(seen).toEqual(['open:m2', 'unpin:m1']);
	});
});

describe('message toolbar pin action', () => {
	it('offers Pin or Unpin by pinned state and marks pinned messages', async () => {
		const list = document.createElement('office-ui-chat-list') as OfficeUiChatList;
		list.selfId = 'ada';
		list.messages = [
			{ id: 'm1', authorId: 'bob', authorName: 'Bob', text: 'one', ts: 1 },
			{ id: 'm2', authorId: 'bob', authorName: 'Bob', text: 'two', ts: 2 },
		];
		list.pinnedIds = ['m2'];
		document.body.append(list);
		await list.updateComplete;
		const events: string[] = [];
		for (const name of ['office-chat-pin', 'office-chat-unpin'])
			list.addEventListener(name, (e) =>
				events.push(`${name}:${(e as CustomEvent).detail.messageId}`),
			);
		const rows = list.shadowRoot!.querySelectorAll<HTMLElement>('.msg');
		rows[0]!.querySelector<HTMLButtonElement>('[aria-label="Pin"]')!.click();
		rows[1]!.querySelector<HTMLButtonElement>('[aria-label="Unpin"]')!.click();
		expect(events).toEqual(['office-chat-pin:m1', 'office-chat-unpin:m2']);
		expect(rows[1]!.querySelector('.pin-mark')).not.toBeNull();
		expect(rows[0]!.querySelector('.pin-mark')).toBeNull();
	});
});
