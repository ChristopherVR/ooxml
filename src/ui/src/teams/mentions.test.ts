import { type Mention, type MentionTarget, resolveMentions } from 'ooxml-core/teams';
import { registerOfficeUi } from '../index';
import type { OfficeUiChatComposer } from './chat-composer';
import type { ChatMessage, OfficeUiChatList } from './chat-list';

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

const ada: MentionTarget = { kind: 'user', id: 'ada', name: 'Ada Lovelace' };
const bob: MentionTarget = { kind: 'user', id: 'bob', name: 'Bob' };
const everyone: MentionTarget = { kind: 'channel', id: 'channel', name: 'channel' };

describe('the @ picker in office-ui-chat-composer', () => {
	const setup = async () => {
		const el = await mount<OfficeUiChatComposer>('office-ui-chat-composer', {
			mentionTargets: [ada, bob, everyone],
		});
		const area = q<HTMLTextAreaElement>(el, 'textarea');
		const type = async (value: string): Promise<void> => {
			area.value = value;
			area.setSelectionRange(value.length, value.length);
			area.dispatchEvent(new Event('input'));
			await el.updateComplete;
		};
		const key = async (name: string): Promise<void> => {
			area.dispatchEvent(new KeyboardEvent('keydown', { key: name, cancelable: true }));
			await el.updateComplete;
		};
		const names = (): string[] =>
			qa(el, '.mentions [role="option"] .name').map((n) => n.textContent!);
		return { el, area, type, key, names };
	};

	it('opens on @, filters as you type and closes when nobody matches', async () => {
		const { el, type, names } = await setup();
		expect(el.shadowRoot!.querySelector('.mentions')).toBeNull();
		await type('hi @');
		expect(names()).toEqual(['Ada Lovelace', 'Bob', 'channel']);
		await type('hi @lo');
		expect(names()).toEqual(['Ada Lovelace']);
		await type('hi @zz');
		expect(el.shadowRoot!.querySelector('.mentions')).toBeNull();
		await type('mail me at a@b');
		expect(el.shadowRoot!.querySelector('.mentions')).toBeNull();
	});

	it('moves with the arrow keys and inserts with Enter instead of sending', async () => {
		const { el, area, type, key } = await setup();
		const sent: unknown[] = [];
		el.addEventListener('office-chat-send', (e) => sent.push((e as CustomEvent).detail));
		await type('hi @');
		await key('ArrowDown');
		expect(qa(el, '.mentions [role="option"]')[1]?.getAttribute('aria-selected')).toBe('true');
		expect(area.getAttribute('aria-activedescendant')).toBe('mention-option-1');
		await key('Enter');
		expect(sent).toHaveLength(0);
		expect(el.value).toBe('hi @Bob ');
		expect(el.shadowRoot!.querySelector('.mentions')).toBeNull();
		await key('Enter');
		expect(sent).toEqual([{ text: 'hi @Bob ', files: [], mentions: [bob] }]);
	});

	it('closes with Escape without cancelling the compose and inserts on click', async () => {
		const { el, type, key } = await setup();
		const cancelled: unknown[] = [];
		el.addEventListener('office-chat-cancel', (e) => cancelled.push(e));
		await type('@Ad');
		await key('Escape');
		expect(el.shadowRoot!.querySelector('.mentions')).toBeNull();
		expect(cancelled).toHaveLength(0);
		await type('@Ad');
		qa<HTMLElement>(el, '.mentions [role="option"]')[0]!.click();
		await el.updateComplete;
		expect(el.value).toBe('@Ada Lovelace ');
	});

	it('offers @channel with a hint', async () => {
		const { el, type } = await setup();
		await type('@chan');
		expect(q(el, '.mentions .hint').textContent).toBe('Notify everyone in this channel');
	});
});

describe('mentions in office-ui-chat-list', () => {
	const message = (text: string, targets: MentionTarget[]): ChatMessage => {
		const mentions: Mention[] = resolveMentions(text, targets);
		return { id: 'm1', authorId: 'x', authorName: 'Xena', text, ts: Date.now(), mentions };
	};

	it('highlights names, emphasises mentions of you and keeps the text exact', async () => {
		const el = await mount<OfficeUiChatList>('office-ui-chat-list', {
			selfId: 'ada',
			messages: [message('hi @Ada Lovelace and @Bob!', [ada, bob])],
		});
		const marks = qa<HTMLElement>(el, '.mention');
		expect(marks.map((m) => [m.textContent, m.dataset.self])).toEqual([
			['@Ada Lovelace', 'true'],
			['@Bob', 'false'],
		]);
		expect(q(el, '.text').textContent).toBe('hi @Ada Lovelace and @Bob!');
		expect(q<HTMLElement>(el, '.msg').dataset.mentionsYou).toBe('true');
	});

	it('does not flag a message that mentions somebody else, or plain messages', async () => {
		const el = await mount<OfficeUiChatList>('office-ui-chat-list', {
			selfId: 'zed',
			messages: [message('@Bob look', [bob]), { ...message('plain', []), id: 'm2' }],
		});
		expect(qa<HTMLElement>(el, '.msg').map((m) => m.dataset.mentionsYou)).toEqual([
			'false',
			'false',
		]);
		expect(qa(el, '.mention[data-self="true"]')).toHaveLength(0);
	});

	it('treats @channel as addressed to everyone', async () => {
		const el = await mount<OfficeUiChatList>('office-ui-chat-list', {
			selfId: 'zed',
			messages: [message('@channel standup', [everyone])],
		});
		expect(q<HTMLElement>(el, '.mention').dataset.self).toBe('true');
		expect(q<HTMLElement>(el, '.msg').dataset.mentionsYou).toBe('true');
	});
});
