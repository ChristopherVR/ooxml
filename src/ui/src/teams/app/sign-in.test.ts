import { parseServerConfig } from 'ooxml-core/teams';
import { SignInFlow } from './sign-in-flow';
import { defineTeamsApp, type TeamsApp } from './teams-app';
import { defineTeamsSignIn, type TeamsSignIn } from './teams-sign-in';

const config = parseServerConfig({
	mode: 'server',
	syncUrl: 'wss://chat.example.com/sync',
	token: 'shared',
}).config;
const SECRET = 'correct horse battery staple';

interface Call {
	url: string;
	body: Record<string, unknown> | null;
}
const calls: Call[] = [];
type Answer = { status: number; body?: unknown; headers?: Record<string, string> } | 'down';
let answers: Record<string, Answer> = {};

function stubServer(next: Record<string, Answer>): void {
	answers = next;
	calls.length = 0;
	vi.stubGlobal(
		'fetch',
		vi.fn(async (url: string, init?: RequestInit) => {
			const path = new URL(url).pathname;
			calls.push({ url, body: init?.body ? JSON.parse(String(init.body)) : null });
			const answer = answers[path] ?? { status: 404 };
			if (answer === 'down') throw new TypeError('network down');
			const body = answer.body === undefined ? null : JSON.stringify(answer.body);
			return new Response(answer.status === 204 ? null : body, {
				status: answer.status,
				...(answer.headers ? { headers: answer.headers } : {}),
			});
		}),
	);
}
const session = (user = 'alice', expiresAt = Date.now() + 60_000) => ({
	status: 200,
	body: { session: 'tok.en.sig', user, expiresAt },
});
const settle = async (): Promise<void> => {
	for (let i = 0; i < 6; i++) await Promise.resolve();
};

afterEach(() => {
	vi.useRealTimers();
	vi.unstubAllGlobals();
	document.body.replaceChildren();
	localStorage.clear();
});

function flow() {
	const changed = vi.fn();
	const events: [string, unknown][] = [];
	const f = new SignInFlow({ changed, emit: (name, detail) => events.push([name, detail]) });
	return { f, changed, events };
}

describe('SignInFlow: finding out whether a sign-in is needed', () => {
	it('never gates a local workspace or a host-managed session', () => {
		stubServer({});
		const { f } = flow();
		expect(f.gate({ mode: 'local', iceServers: [] })).toBe(false);
		expect(f.gate({ ...config, sessionToken: 'from-host' })).toBe(false);
		expect(calls).toHaveLength(0);
	});

	it('asks /health once and shows the sign-in when accounts are required', async () => {
		stubServer({ '/health': { status: 200, body: { ok: true, accounts: true } } });
		const { f, changed } = flow();
		expect(f.gate(config)).toBe(true);
		expect(f.phase).toBe('checking');
		expect(f.gate(config)).toBe(true);
		await settle();
		expect(calls.filter((c) => c.url.endsWith('/health'))).toHaveLength(1);
		expect(f.phase).toBe('required');
		expect(changed).toHaveBeenCalled();
		expect(f.gate(config)).toBe(true);
	});

	it('leaves the zero-config dev server alone', async () => {
		stubServer({ '/health': { status: 200, body: { ok: true, accounts: false } } });
		const { f } = flow();
		f.gate(config);
		await settle();
		expect(f.phase).toBe('open');
		expect(f.gate(config)).toBe(false);
	});

	it('carries on when the server cannot be asked, and asks again next time', async () => {
		stubServer({ '/health': 'down' });
		const { f } = flow();
		f.gate(config);
		await settle();
		expect(f.phase).toBe('open');
		stubServer({ '/health': { status: 200, body: { ok: true, accounts: true } } });
		expect(f.gate(config)).toBe(true);
		await settle();
		expect(f.phase).toBe('required');
	});
});

describe('SignInFlow: signing in', () => {
	async function required() {
		stubServer({
			'/health': { status: 200, body: { accounts: true } },
			'/auth/session': session(),
		});
		const made = flow();
		made.f.gate(config);
		await settle();
		return made;
	}

	it('stores the session, announces it and starts the client with it', async () => {
		const { f, events } = await required();
		await f.submit(config, ' alice ', SECRET);
		expect(f.session).toMatchObject({ userId: 'alice', token: 'tok.en.sig' });
		expect(f.phase).toBe('open');
		expect(f.gate(config)).toBe(false);
		expect(f.apply(config).sessionToken).toBe('tok.en.sig');
		expect(events[0]).toEqual(['teams-signed-in', expect.objectContaining({ userId: 'alice' })]);
		const post = calls.find((c) => c.url.endsWith('/auth/session'));
		expect(post?.body).toEqual({ user: 'alice', credential: SECRET });
	});

	it('explains a wrong password without saying which field was wrong', async () => {
		const { f } = await required();
		answers['/auth/session'] = { status: 401 };
		await f.submit(config, 'alice', SECRET);
		expect(f.session).toBeNull();
		expect(f.error).toBe('The user ID or password is not correct.');
		expect(f.userId).toBe('alice');
		expect(f.busy).toBe(false);
		expect(f.error).not.toContain(SECRET);
	});

	it('passes the rate-limit wait on', async () => {
		const { f } = await required();
		answers['/auth/session'] = { status: 429, headers: { 'Retry-After': '30' } };
		await f.submit(config, 'alice', SECRET);
		expect(f.error).toBe('Too many attempts. Try again in 30 seconds.');
		expect(f.retryAfterSec).toBe(30);
	});

	it('reports an unreachable server and a server error', async () => {
		const { f } = await required();
		answers['/auth/session'] = 'down';
		await f.submit(config, 'alice', SECRET);
		expect(f.error).toContain('Could not reach the server');
		answers['/auth/session'] = { status: 500 };
		await f.submit(config, 'alice', SECRET);
		expect(f.error).toContain('could not sign you in');
	});

	it('ignores a second submit while one is running', async () => {
		const { f } = await required();
		const first = f.submit(config, 'alice', SECRET);
		await f.submit(config, 'alice', SECRET);
		await first;
		expect(calls.filter((c) => c.url.endsWith('/auth/session'))).toHaveLength(1);
	});
});

describe('SignInFlow: a session that stops working', () => {
	async function signedIn(expiresAt = Date.now() + 60_000) {
		vi.useFakeTimers();
		stubServer({
			'/health': { status: 200, body: { accounts: true } },
			'/auth/session': session('alice', expiresAt),
			'/auth/check': { status: 204 },
			'/auth/signout': { status: 204 },
		});
		const made = flow();
		made.f.gate(config);
		await vi.advanceTimersByTimeAsync(0);
		await made.f.submit(config, 'alice', SECRET);
		return made;
	}

	it('asks for a new sign-in when the session expires, keeping the user id', async () => {
		const { f, events } = await signedIn(Date.now() + 10_000);
		expect(f.session).not.toBeNull();
		await vi.advanceTimersByTimeAsync(10_001);
		expect(f.session).toBeNull();
		expect(f.phase).toBe('required');
		expect(f.userId).toBe('alice');
		expect(f.notice).toContain('expired');
		expect(events.map((e) => e[0])).toContain('teams-session-ended');
		expect(f.gate(config)).toBe(true);
	});

	it('asks for a new sign-in when the server rejects the session', async () => {
		const { f } = await signedIn();
		f.watchConnection(false, config);
		answers['/auth/check'] = { status: 401 };
		await vi.advanceTimersByTimeAsync(4001);
		expect(f.session).toBeNull();
		expect(f.notice).toContain('ended your session');
		const check = calls.find((c) => c.url.endsWith('/auth/check'));
		expect(check?.body).toEqual({ session: 'tok.en.sig' });
	});

	it('keeps the session while the server is merely unreachable, and checks again', async () => {
		const { f } = await signedIn();
		answers['/auth/check'] = 'down';
		f.watchConnection(false, config);
		await vi.advanceTimersByTimeAsync(4001);
		expect(f.session).not.toBeNull();
		answers['/auth/check'] = { status: 401 };
		await vi.advanceTimersByTimeAsync(15_001);
		expect(f.session).toBeNull();
	});

	it('does not check a session while the connection is fine', async () => {
		const { f } = await signedIn();
		f.watchConnection(false, config);
		f.watchConnection(true, config);
		await vi.advanceTimersByTimeAsync(60_000 - 1);
		expect(calls.some((c) => c.url.endsWith('/auth/check'))).toBe(false);
	});

	it('signs out locally at once and on the server afterwards', async () => {
		const { f, events } = await signedIn();
		await f.signOut(config);
		expect(f.session).toBeNull();
		expect(f.phase).toBe('required');
		expect(f.userId).toBe('');
		expect(f.notice).toBe('You signed out.');
		expect(events.at(-1)?.[0]).toBe('teams-signed-out');
		const out = calls.find((c) => c.url.endsWith('/auth/signout'));
		expect(out?.body).toEqual({ session: 'tok.en.sig' });
		expect(f.apply(config).sessionToken).toBeUndefined();
	});

	it('stops its timers when disposed', async () => {
		const { f } = await signedIn(Date.now() + 10_000);
		f.dispose();
		await vi.advanceTimersByTimeAsync(20_000);
		expect(f.session).not.toBeNull();
	});
});

describe('<teams-sign-in>', () => {
	beforeAll(() => defineTeamsSignIn());

	async function mount(props: Partial<TeamsSignIn> = {}): Promise<TeamsSignIn> {
		const el = document.createElement('teams-sign-in') as TeamsSignIn;
		Object.assign(el, props);
		document.body.append(el);
		await el.updateComplete;
		return el;
	}
	const fill = (el: TeamsSignIn, user: string, credential: string): HTMLFormElement => {
		const form = el.shadowRoot!.querySelector('form')!;
		form.querySelector<HTMLInputElement>('input[name=user]')!.value = user;
		form.querySelector<HTMLInputElement>('input[name=credential]')!.value = credential;
		return form;
	};

	it('has a labelled user id and a password field', async () => {
		const el = await mount({ userId: 'alice' });
		const user = el.shadowRoot!.querySelector<HTMLInputElement>('input[name=user]')!;
		const secret = el.shadowRoot!.querySelector<HTMLInputElement>('input[name=credential]')!;
		expect(user.value).toBe('alice');
		expect(user.autocomplete).toBe('username');
		expect(secret.type).toBe('password');
		expect(secret.autocomplete).toBe('current-password');
		expect(el.shadowRoot!.textContent).toContain('User ID');
		expect(el.shadowRoot!.textContent).toContain('Password');
	});

	it('emits the credentials once and clears the password field', async () => {
		const el = await mount();
		const seen: unknown[] = [];
		el.addEventListener('teams-sign-in-submit', (e) => seen.push((e as CustomEvent).detail));
		const form = fill(el, ' alice ', SECRET);
		form.dispatchEvent(new Event('submit', { cancelable: true }));
		expect(seen).toEqual([{ userId: 'alice', credential: SECRET }]);
		expect(form.querySelector<HTMLInputElement>('input[name=credential]')!.value).toBe('');
	});

	it('does not emit an empty form or while busy', async () => {
		const el = await mount();
		const seen: unknown[] = [];
		el.addEventListener('teams-sign-in-submit', (e) => seen.push(e));
		fill(el, 'alice', '').dispatchEvent(new Event('submit', { cancelable: true }));
		el.busy = true;
		await el.updateComplete;
		fill(el, 'alice', SECRET).dispatchEvent(new Event('submit', { cancelable: true }));
		expect(seen).toHaveLength(0);
	});

	it('announces an error and the reason it is back, never the password', async () => {
		const el = await mount({
			error: 'The user ID or password is not correct.',
			notice: 'Your session expired.',
		});
		const alert = el.shadowRoot!.querySelector('[role=alert]')!;
		expect(alert.textContent).toContain('not correct');
		expect(el.shadowRoot!.textContent).toContain('Your session expired.');
		expect(el.shadowRoot!.innerHTML).not.toContain(SECRET);
	});

	it('disables the button and counts down a rate-limit wait', async () => {
		vi.useFakeTimers();
		const el = await mount({ error: 'Too many attempts. Try again in 3 seconds.', retryAfter: 3 });
		const button = (): HTMLButtonElement => el.shadowRoot!.querySelector('button')!;
		expect(button().disabled).toBe(true);
		await vi.advanceTimersByTimeAsync(1000);
		await el.updateComplete;
		expect(el.shadowRoot!.querySelector('[role=alert]')!.textContent).toContain('2 seconds');
		await vi.advanceTimersByTimeAsync(2000);
		await el.updateComplete;
		expect(button().disabled).toBe(false);
		el.remove();
	});

	it('shows a status while the server is being asked', async () => {
		const el = await mount({ checking: true });
		expect(el.shadowRoot!.querySelector('form')).toBeNull();
		expect(el.shadowRoot!.querySelector('[role=status]')).not.toBeNull();
	});
});

describe('<teams-app> with a server that requires accounts', () => {
	beforeAll(() => defineTeamsApp());
	const SERVER = JSON.stringify({ mode: 'server', syncUrl: 'ws://127.0.0.1:9/sync' });

	async function mount(attrs: Record<string, string>): Promise<TeamsApp> {
		const el = document.createElement('teams-app') as TeamsApp;
		for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, v);
		document.body.append(el);
		await el.updateComplete;
		return el;
	}
	const until = async (el: TeamsApp, check: () => boolean): Promise<void> => {
		for (let i = 0; i < 40 && !check(); i++) {
			await new Promise((r) => setTimeout(r, 10));
			await el.updateComplete;
		}
	};

	it('shows the sign-in instead of the welcome screen and signs in', async () => {
		stubServer({
			'/health': { status: 200, body: { accounts: true } },
			'/auth/session': session('alice'),
			'/auth/signout': { status: 204 },
		});
		const el = await mount({ 'workspace-id': 'si-app', 'server-config': SERVER });
		await until(
			el,
			() => el.shadowRoot!.querySelector<TeamsSignIn>('teams-sign-in')?.checking === false,
		);
		const form = el.shadowRoot!.querySelector<TeamsSignIn>('teams-sign-in')!;
		expect(el.shadowRoot!.querySelector('.welcome')).toBeNull();
		expect(el.client).toBeNull();
		await form.updateComplete;
		const inner = form.shadowRoot!.querySelector('form')!;
		inner.querySelector<HTMLInputElement>('input[name=user]')!.value = 'alice';
		inner.querySelector<HTMLInputElement>('input[name=credential]')!.value = SECRET;
		const signedIn = vi.fn();
		el.addEventListener('teams-signed-in', signedIn);
		inner.dispatchEvent(new Event('submit', { cancelable: true }));
		await until(el, () => el.client !== null);
		expect(signedIn).toHaveBeenCalledTimes(1);
		expect(el.shadowRoot!.querySelector('teams-sign-in')).toBeNull();
		expect(el.client?.getState().user.id).toBe('alice');
		expect(el.shadowRoot!.querySelector('.shell')).not.toBeNull();
		// The profile menu names the account and offers Sign out.
		const menu = el.shadowRoot!.querySelector('teams-profile-menu') as HTMLElement & {
			updateComplete: Promise<boolean>;
		};
		await menu.updateComplete;
		expect(menu.shadowRoot!.textContent).toContain('Signed in as alice');
		menu.shadowRoot!.querySelector<HTMLButtonElement>('button[name=sign-out]')!.click();
		await until(el, () => el.client === null);
		expect(el.shadowRoot!.querySelector('teams-sign-in')).not.toBeNull();
		expect(calls.some((c) => c.url.endsWith('/auth/signout'))).toBe(true);
		el.remove();
	});

	it('leaves the dev server without accounts as it was', async () => {
		stubServer({ '/health': { status: 200, body: { ok: true, accounts: false } } });
		const el = await mount({
			'workspace-id': 'si-dev',
			'user-name': 'Ada',
			'user-id': 'ada',
			'server-config': SERVER,
		});
		await until(el, () => el.client !== null);
		expect(el.shadowRoot!.querySelector('teams-sign-in')).toBeNull();
		expect(el.client?.getState().user.id).toBe('ada');
		expect(calls.some((c) => c.url.endsWith('/auth/session'))).toBe(false);
		el.remove();
	});

	it('does not ask a local workspace anything', async () => {
		stubServer({});
		const el = await mount({
			'workspace-id': 'si-local',
			'user-name': 'Ada',
			'server-config': JSON.stringify({ mode: 'local', iceServers: [] }),
		});
		expect(el.client).not.toBeNull();
		expect(calls).toHaveLength(0);
		el.remove();
	});
});
