// @vitest-environment jsdom
import { defineTeamsApp, type TeamsApp } from './teams-app';

const LOCAL = JSON.stringify({ mode: 'local', iceServers: [] });
const tick = (ms = 0): Promise<void> => new Promise((r) => setTimeout(r, ms));

beforeAll(() => defineTeamsApp());
afterEach(() => {
	document.body.replaceChildren();
	localStorage.clear();
});

async function workspace(id: string) {
	const el = document.createElement('teams-app') as TeamsApp;
	for (const [name, value] of Object.entries({
		'workspace-id': id,
		'user-name': 'Ada',
		'user-id': 'ada',
		'server-config': LOCAL,
	}))
		el.setAttribute(name, value);
	document.body.append(el);
	await el.updateComplete;
	el.client!.createChannel('Ops');
	await tick(10);
	await el.updateComplete;
	const root = el.shadowRoot!;
	const deepActive = (): Element | null => {
		let active = root.activeElement;
		while (active?.shadowRoot?.activeElement) active = active.shadowRoot.activeElement;
		return active;
	};
	/** The region the keyboard is in now, across shadow roots and slots. */
	const inRegion = (selector: string): boolean => {
		const region = root.querySelector(selector);
		for (let node: Node | null = deepActive(); node;)
			if (node === region) return true;
			else
				node =
					(node as Element).assignedSlot ?? node.parentNode ?? (node as ShadowRoot).host ?? null;
		return false;
	};
	const press = (shiftKey = false) =>
		deepActive()?.dispatchEvent(
			new KeyboardEvent('keydown', { key: 'F6', shiftKey, bubbles: true, composed: true }),
		);
	const message = root
		.querySelector('office-ui-chat-composer')!
		.shadowRoot!.querySelector<HTMLTextAreaElement>('textarea')!;
	return { el, root, message, deepActive, inRegion, press };
}

describe('Teams region keys', () => {
	it('walks F6 from the composer round the workspace and back', async () => {
		const { message, deepActive, inRegion, press } = await workspace('tr-forward');
		message.focus();
		press();
		expect(inRegion('office-ui-app-rail')).toBe(true);
		press();
		expect(inRegion('office-ui-channel-list')).toBe(true);
		press();
		expect(inRegion('header.channel-head')).toBe(true);
		press();
		expect(inRegion('office-ui-chat-list')).toBe(true);
		press();
		expect(inRegion('office-ui-chat-composer')).toBe(true);
		// The composer hands the keyboard to the message box, not to Attach or Send.
		expect(deepActive()).toBe(message);
	});

	it('walks Shift+F6 the other way', async () => {
		const { message, inRegion, press } = await workspace('tr-back');
		message.focus();
		press(true);
		expect(inRegion('office-ui-chat-list')).toBe(true);
		press(true);
		expect(inRegion('header.channel-head')).toBe(true);
		press(true);
		expect(inRegion('office-ui-channel-list')).toBe(true);
		press(true);
		expect(inRegion('office-ui-app-rail')).toBe(true);
		press(true);
		expect(inRegion('office-ui-chat-composer')).toBe(true);
	});
});
