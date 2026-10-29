import { createChatStore } from 'ooxml-core/teams';
import { defineTeamsApp, type TeamsApp } from './teams-app';
import type { TeamsSignIn } from './teams-sign-in';

// A mounted <teams-app> signed in through a faked server: a moderator session shows the moderation
// actions (Hide on other people's messages, channel options on their channels), a plain user's
// session shows neither.
const SERVER = JSON.stringify({ mode: 'server', syncUrl: 'ws://127.0.0.1:9/sync' });
const SECRET = 'correct horse battery staple';

function stubServer(user: string, role?: 'moderator'): void {
	vi.stubGlobal(
		'fetch',
		vi.fn(async (url: string) => {
			const path = new URL(url).pathname;
			const body =
				path === '/health'
					? { accounts: true }
					: path === '/auth/session'
						? {
								session: 'tok.en.sig',
								user,
								...(role ? { role } : {}),
								expiresAt: Date.now() + 60_000,
							}
						: null;
			return new Response(body ? JSON.stringify(body) : null, { status: body ? 200 : 204 });
		}),
	);
}

beforeAll(() => defineTeamsApp());
afterEach(() => {
	vi.unstubAllGlobals();
	document.body.replaceChildren();
	localStorage.clear();
});

const tick = (ms = 10): Promise<void> => new Promise((r) => setTimeout(r, ms));
async function until(el: TeamsApp, check: () => boolean): Promise<void> {
	for (let i = 0; i < 80 && !check(); i++) {
		await tick();
		await el.updateComplete;
	}
}

let workspace = 0;
async function signedIn(user: string, role?: 'moderator'): Promise<TeamsApp> {
	stubServer(user, role);
	const el = document.createElement('teams-app') as TeamsApp;
	el.setAttribute('workspace-id', `mod-app-${workspace++}`);
	el.setAttribute('server-config', SERVER);
	document.body.append(el);
	await el.updateComplete;
	await until(
		el,
		() => el.shadowRoot!.querySelector<TeamsSignIn>('teams-sign-in')?.checking === false,
	);
	const form = el.shadowRoot!.querySelector<TeamsSignIn>('teams-sign-in')!;
	await form.updateComplete;
	const inner = form.shadowRoot!.querySelector('form')!;
	inner.querySelector<HTMLInputElement>('input[name=user]')!.value = user;
	inner.querySelector<HTMLInputElement>('input[name=credential]')!.value = SECRET;
	inner.dispatchEvent(new Event('submit', { cancelable: true }));
	await until(el, () => el.client !== null);
	return el;
}

/** A channel and a message written by somebody else, as a colleague's client would. */
function seedOthers(el: TeamsApp): string {
	const doc = el.client!.workspace.session.doc;
	const other = createChatStore(doc, { id: 'zed', name: 'Zed' });
	const channel = other.createChannel({ name: 'Zed room' })!;
	other.post(channel.id, { text: 'hello from Zed' });
	el.client!.select(channel.id);
	return channel.id;
}

describe('<teams-app> moderation actions', () => {
	it('shows Hide and the channel options to a moderator session', async () => {
		const el = await signedIn('root', 'moderator');
		expect(el.client!.getState().moderator).toBe(true);
		seedOthers(el);
		await until(el, () => el.shadowRoot!.querySelector('[data-channel-options]') !== null);
		expect(el.shadowRoot!.querySelector('[data-channel-options]')).not.toBeNull();
		const list = el.shadowRoot!.querySelector('office-ui-chat-list') as HTMLElement & {
			updateComplete: Promise<boolean>;
		};
		await until(el, () => list.shadowRoot?.querySelector('.toolbar') != null);
		await list.updateComplete;
		expect(list.shadowRoot!.querySelector('[data-action="hide"]')).not.toBeNull();
		el.remove();
	});

	it('shows neither to a plain user, who does not own that channel or message', async () => {
		const el = await signedIn('bob');
		expect(el.client!.getState().moderator).toBe(false);
		seedOthers(el);
		const list = el.shadowRoot!.querySelector('office-ui-chat-list') as HTMLElement & {
			updateComplete: Promise<boolean>;
		};
		await until(el, () => list.shadowRoot?.querySelector('.toolbar') != null);
		await list.updateComplete;
		expect(list.shadowRoot!.querySelector('.toolbar')).not.toBeNull();
		expect(list.shadowRoot!.querySelector('[data-action="hide"]')).toBeNull();
		expect(el.shadowRoot!.querySelector('[data-channel-options]')).toBeNull();
		el.remove();
	});
});
