import type { ChatView } from 'ooxml-core/teams';
import { registerOfficeUi } from '../index';
import type { OfficeUiChatSidebar } from './chat-sidebar';

beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

const chat = (over: Partial<ChatView> & { id: string }): ChatView => ({
	kind: 'direct',
	title: 'Bob',
	name: '',
	members: [
		{ id: 'ada', name: 'Ada', self: true, availability: 'available' },
		{ id: 'bob', name: 'Bob', self: false, availability: 'busy' },
	],
	unread: 0,
	lastMessage: {
		authorName: 'Bob',
		text: 'see you soon',
		ts: Date.now(),
		mine: false,
		system: false,
	},
	ts: Date.now(),
	live: false,
	canLeave: false,
	...over,
});

async function mount(chats: ChatView[], selectedId = ''): Promise<OfficeUiChatSidebar> {
	const el = document.createElement('office-ui-chat-sidebar') as OfficeUiChatSidebar;
	el.chats = chats;
	el.selectedId = selectedId;
	document.body.append(el);
	await el.updateComplete;
	return el;
}
const qa = <T extends Element>(el: HTMLElement, s: string): T[] => [
	...el.shadowRoot!.querySelectorAll<T>(s),
];

describe('office-ui-chat-sidebar', () => {
	it('shows presence on 1:1 avatars, the preview, unread badges and the selection', async () => {
		const el = await mount(
			[
				chat({ id: 'a', unread: 3 }),
				chat({
					id: 'g',
					kind: 'group',
					title: 'Launch',
					lastMessage: {
						authorName: 'Cy',
						text: 'ship it',
						ts: Date.now(),
						mine: false,
						system: false,
					},
				}),
			],
			'g',
		);
		const rows = qa<HTMLElement>(el, '.row');
		expect(rows.map((r) => r.querySelector('.title')!.textContent)).toEqual(['Bob', 'Launch']);
		expect(rows[0]!.querySelector('office-ui-avatar')!.getAttribute('presence')).toBe('busy');
		expect(rows[1]!.querySelector('office-ui-avatar')).toBeNull();
		expect(rows[0]!.querySelector('.preview')!.textContent).toBe('Bob: see you soon');
		expect(rows[0]!.querySelector('.count')!.textContent).toBe('3');
		expect(rows[0]!.dataset.unread).toBe('true');
		expect(rows[1]!.getAttribute('aria-current')).toBe('true');
	});

	it('words your own last message and system notices', async () => {
		const el = await mount([
			chat({
				id: 'a',
				lastMessage: { authorName: 'Ada', text: 'on it', ts: 1, mine: true, system: false },
			}),
			chat({
				id: 'b',
				lastMessage: {
					authorName: 'Ada',
					text: 'Dan joined the chat',
					ts: 1,
					mine: false,
					system: true,
				},
			}),
			chat({ id: 'c', lastMessage: null }),
		]);
		expect(qa(el, '.preview').map((p) => p.textContent)).toEqual([
			'You: on it',
			'Dan joined the chat',
			'No messages yet',
		]);
	});

	it('emits select and new chat events', async () => {
		const el = await mount([chat({ id: 'a' })]);
		const got: unknown[] = [];
		el.addEventListener('office-chat-select', (e) => got.push((e as CustomEvent).detail));
		el.addEventListener('office-chat-new', () => got.push('new'));
		qa<HTMLButtonElement>(el, '.row')[0]!.click();
		el.shadowRoot!.querySelector<HTMLButtonElement>('.new')!.click();
		expect(got).toEqual([{ id: 'a' }, 'new']);
	});

	it('explains an empty list', async () => {
		const el = await mount([]);
		expect(el.shadowRoot!.querySelector('.empty')!.textContent).toContain('No chats yet');
	});
});
