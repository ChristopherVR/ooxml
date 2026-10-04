import { registerOfficeUi } from '../index.js';
import type { OfficeUiCallControls } from './call-controls.js';
import type { OfficeUiCallGrid } from './call-grid.js';
import type { OfficeUiChannelList } from './channel-list.js';
import type { OfficeUiChatComposer } from './chat-composer.js';
import type { ChatMessage, OfficeUiChatList } from './chat-list.js';
import { initialsOf } from '../presence.js';
import { dayLabel, formatSize } from './base.js';
import { formatElapsed } from './call-controls.js';

beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

async function mount<T extends HTMLElement & { updateComplete: Promise<boolean> }>(
	tag: string,
	props: Partial<T> = {},
): Promise<T> {
	const el = document.createElement(tag) as T;
	Object.assign(el, props);
	document.body.append(el);
	await el.updateComplete;
	return el;
}
const q = <T extends Element>(el: HTMLElement, selector: string): T =>
	el.shadowRoot!.querySelector<T>(selector)!;
const qa = <T extends Element>(el: HTMLElement, selector: string): T[] => [
	...el.shadowRoot!.querySelectorAll<T>(selector),
];
const events = (el: HTMLElement, type: string): unknown[] => {
	const out: unknown[] = [];
	el.addEventListener(type, (e) => out.push((e as CustomEvent).detail));
	return out;
};

const msg = (over: Partial<ChatMessage> & { id: string }): ChatMessage => ({
	authorId: 'x',
	authorName: 'Xena',
	text: 'hi',
	ts: Date.now(),
	...over,
});

describe('helpers', () => {
	it('formats initials, sizes, elapsed time and day labels', () => {
		expect(initialsOf('Ada Lovelace')).toBe('AL');
		expect(initialsOf('  ')).toBe('?');
		expect(formatSize(512)).toBe('512 B');
		expect(formatSize(2048)).toBe('2 KB');
		expect(formatElapsed(65_000)).toBe('01:05');
		expect(formatElapsed(3_725_000)).toBe('1:02:05');
		expect(dayLabel(Date.now())).toBe('Today');
		expect(dayLabel(Date.now() - 86_400_000)).toBe('Yesterday');
	});
});

describe('office-ui-avatar', () => {
	it('names itself with initials and presence for assistive tech', async () => {
		const el = await mount('office-ui-avatar', { name: 'Grace Hopper', presence: 'busy' } as never);
		expect(q(el, '.face').textContent).toBe('GH');
		expect(el.getAttribute('aria-label')).toBe('Grace Hopper, Busy');
		expect(q(el, '.presence').getAttribute('data-presence')).toBe('busy');
	});
});

describe('office-ui-chat-list', () => {
	it('renders peer text as text, never as markup', async () => {
		const el = await mount<OfficeUiChatList>('office-ui-chat-list', {
			messages: [msg({ id: '1', text: '<img src=x onerror=alert(1)>' })],
		});
		expect(el.shadowRoot!.querySelector('img')).toBeNull();
		expect(q(el, '.text').textContent).toContain('<img');
	});

	it('groups close messages, separates days and quotes replies', async () => {
		const now = Date.now();
		const el = await mount<OfficeUiChatList>('office-ui-chat-list', {
			messages: [
				msg({ id: '1', ts: now - 3 * 86_400_000 }),
				msg({ id: '2', ts: now - 60_000, text: 'first' }),
				msg({ id: '3', ts: now - 30_000, text: 'second' }),
				msg({
					id: '4',
					ts: now - 10_000,
					authorId: 'y',
					authorName: 'Yan',
					replyTo: '2',
					text: 'answer',
				}),
			],
		});
		expect(qa(el, '.day')).toHaveLength(2);
		const articles = qa<HTMLElement>(el, 'article');
		expect(articles.map((a) => a.dataset.grouped)).toEqual(['false', 'false', 'true', 'false']);
		expect(articles[3]!.querySelector('blockquote')!.textContent).toContain('first');
	});

	it('offers edit and delete only on own messages and toggles reactions by event', async () => {
		const el = await mount<OfficeUiChatList>('office-ui-chat-list', {
			selfId: 'me',
			messages: [
				msg({ id: '1', authorId: 'me', authorName: 'Me', text: 'mine' }),
				msg({ id: '2', text: 'theirs', reactions: { '👍': ['me'] } }),
			],
		});
		const [mine, theirs] = qa<HTMLElement>(el, 'article');
		expect(mine!.querySelector('[aria-label="Edit"]')).not.toBeNull();
		expect(theirs!.querySelector('[aria-label="Edit"]')).toBeNull();
		const reacted = events(el, 'office-chat-react');
		(theirs!.querySelector('.pill') as HTMLElement).click();
		(theirs!.querySelector('[aria-label="React ❤️"]') as HTMLElement).click();
		expect(reacted).toEqual([
			{ messageId: '2', emoji: '👍' },
			{ messageId: '2', emoji: '❤️' },
		]);
		expect(theirs!.querySelector('.pill')!.getAttribute('aria-pressed')).toBe('true');
	});

	it('shows Office file cards that report which file was opened', async () => {
		const el = await mount<OfficeUiChatList>('office-ui-chat-list', {
			messages: [
				msg({ id: '1', attachments: [{ name: 'Budget.xlsx', kind: 'xlsx', size: 2048 }] }),
			],
		});
		const opened = events(el, 'office-chat-open-file');
		expect(q(el, '.badge').textContent).toBe('X');
		(q(el, '.file') as HTMLElement).click();
		expect(opened).toEqual([{ attachment: { name: 'Budget.xlsx', kind: 'xlsx', size: 2048 } }]);
	});

	it('shows deleted messages as a tombstone without text or toolbar', async () => {
		const el = await mount<OfficeUiChatList>('office-ui-chat-list', {
			messages: [msg({ id: '1', deleted: true, text: '' })],
		});
		expect(q(el, '.deleted').textContent).toBe('This message was deleted');
		expect(el.shadowRoot!.querySelector('.toolbar')).toBeNull();
	});
});

describe('office-ui-chat-composer', () => {
	const type = async (el: OfficeUiChatComposer, value: string): Promise<HTMLTextAreaElement> => {
		const area = q<HTMLTextAreaElement>(el, 'textarea');
		area.value = value;
		area.dispatchEvent(new Event('input'));
		await el.updateComplete;
		return area;
	};

	it('sends on Enter, keeps Shift+Enter for new lines and ignores empty text', async () => {
		const el = await mount<OfficeUiChatComposer>('office-ui-chat-composer');
		const sent = events(el, 'office-chat-send');
		const area = q<HTMLTextAreaElement>(el, 'textarea');
		area.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }));
		expect(sent).toHaveLength(0);
		await type(el, 'hi');
		area.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', shiftKey: true }));
		expect(sent).toHaveLength(0);
		area.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }));
		await el.updateComplete;
		expect(sent).toEqual([{ text: 'hi', files: [] }]);
		expect(area.value).toBe('');
		expect(q<HTMLButtonElement>(el, '.send').disabled).toBe(true);
	});

	it('inserts emoji, shows the reply banner and cancels it with Escape', async () => {
		const el = await mount<OfficeUiChatComposer>('office-ui-chat-composer', { replyingTo: 'Xena' });
		expect(q(el, '.banner').textContent).toContain('Replying to Xena');
		const cancelled = events(el, 'office-chat-cancel');
		q<HTMLTextAreaElement>(el, 'textarea').dispatchEvent(
			new KeyboardEvent('keydown', { key: 'Escape' }),
		);
		expect(cancelled).toHaveLength(1);
		(q(el, '[aria-label="Emoji"]') as HTMLElement).click();
		await el.updateComplete;
		qa<HTMLElement>(el, '.picker button')[0]!.click();
		await el.updateComplete;
		expect(el.value).toBe('😀');
	});

	it('throttles typing notifications and shows who is typing', async () => {
		const el = await mount<OfficeUiChatComposer>('office-ui-chat-composer', {
			typing: ['Ana', 'Bo'],
		});
		const typing = events(el, 'office-chat-typing');
		await type(el, 'a');
		await type(el, 'ab');
		expect(typing).toHaveLength(1);
		expect(q(el, '.typing').textContent).toBe('Ana, Bo are typing…');
	});
});

describe('office-ui-channel-list', () => {
	it('marks the selected and unread channels and emits selection and creation', async () => {
		const el = await mount<OfficeUiChannelList>('office-ui-channel-list', {
			heading: 'Acme',
			selectedId: 'a',
			channels: [
				{ id: 'a', name: 'General' },
				{ id: 'b', name: 'Design', unread: 3, live: true },
			],
		});
		const rows = qa<HTMLElement>(el, '.row');
		expect(rows[0]!.getAttribute('aria-current')).toBe('true');
		expect(rows[1]!.dataset.unread).toBe('true');
		expect(rows[1]!.querySelector('.count')!.textContent).toBe('3');
		expect(rows[1]!.querySelector('.live')).not.toBeNull();
		const selected = events(el, 'office-channel-select');
		const created = events(el, 'office-channel-create');
		rows[1]!.click();
		(q(el, '.add') as HTMLElement).click();
		expect(selected).toEqual([{ id: 'b' }]);
		expect(created).toHaveLength(1);
	});

	it('collapses its channels', async () => {
		const el = await mount<OfficeUiChannelList>('office-ui-channel-list', {
			channels: [{ id: 'a', name: 'A' }],
		});
		(q(el, '.toggle') as HTMLElement).click();
		await el.updateComplete;
		expect(q(el, 'ul').hasAttribute('hidden')).toBe(true);
	});
});

describe('office-ui-call-controls', () => {
	it('reflects state in labels and emits actions', async () => {
		const el = await mount<OfficeUiCallControls>('office-ui-call-controls', {
			mic: true,
			camera: false,
			panel: 'chat',
		});
		const actions = events(el, 'office-call-action');
		expect(q(el, '[data-action="mic"]').getAttribute('aria-label')).toBe('Mute');
		expect(q(el, '[data-action="camera"]').getAttribute('aria-label')).toBe('Camera on');
		expect(q(el, '[data-action="chat"]').getAttribute('aria-pressed')).toBe('true');
		(q(el, '[data-action="mic"]') as HTMLElement).click();
		(q(el, '.leave') as HTMLElement).click();
		expect(actions).toEqual([{ action: 'mic' }, { action: 'leave' }]);
	});
	it('shows elapsed time once a start time is known', async () => {
		const el = await mount<OfficeUiCallControls>('office-ui-call-controls', {
			startedAt: Date.now() - 61_000,
		});
		expect(q(el, 'time').textContent).toMatch(/^01:0[1-2]$/);
		el.remove();
	});
});

describe('office-ui-call-grid', () => {
	it('shows an avatar for tiles without video and spotlights a screen share', async () => {
		const el = await mount<OfficeUiCallGrid>('office-ui-call-grid', {
			participants: [
				{ id: 'a', name: 'Ada Lovelace', self: true },
				{ id: 'b', name: 'Bob', screen: true, audio: false },
			],
		});
		const grid = q<HTMLElement>(el, '.grid');
		expect(grid.dataset.spotlight).toBe('true');
		const tiles = qa<HTMLElement>(el, '.tile');
		expect(tiles[0]!.dataset.main).toBe('true');
		expect(tiles[0]!.querySelector('.name')!.textContent).toContain('Bob');
		expect(tiles[0]!.querySelector('.muted')).not.toBeNull();
		expect(tiles[1]!.querySelector('office-ui-avatar')!.getAttribute('name')).toBe('Ada Lovelace');
	});
	it('says so when nobody is in the call', async () => {
		const el = await mount<OfficeUiCallGrid>('office-ui-call-grid');
		expect(q(el, '.empty').textContent).toBe('Nobody is in the call');
	});
});
