import { applyTeamsProps, listenTeamsEvents } from './bind.js';
import { iceToText, parseIceLines } from './teams-settings.js';
import { defineTeamsApp, TeamsApp } from './teams-app.js';

const LOCAL = JSON.stringify({ mode: 'local', iceServers: [] });
const tick = (ms = 0): Promise<void> => new Promise((r) => setTimeout(r, ms));
const q = <T extends Element>(root: ParentNode | null | undefined, s: string): T => root!.querySelector<T>(s)!;

beforeAll(() => defineTeamsApp());
afterEach(() => {
	document.body.replaceChildren();
	localStorage.clear();
});

async function mount(attrs: Record<string, string> = {}): Promise<TeamsApp> {
	const el = document.createElement('teams-app') as TeamsApp;
	for (const [k, v] of Object.entries(attrs)) el.setAttribute(k, v);
	document.body.append(el);
	await el.updateComplete;
	return el;
}

describe('<teams-app>', () => {
	it('asks for a name first and remembers it', async () => {
		const el = await mount({ 'workspace-id': 'ta-welcome', 'server-config': LOCAL });
		const form = q<HTMLFormElement>(el.shadowRoot, '.welcome form');
		q<HTMLInputElement>(form, 'input').value = 'Grace';
		form.requestSubmit();
		await tick(20);
		await el.updateComplete;
		expect(JSON.parse(localStorage.getItem('teams:identity')!).name).toBe('Grace');
		expect(el.shadowRoot!.querySelector('.shell')).not.toBeNull();
		el.remove();
	});

	it('renders the Teams shell from a named user and ooxml-ui elements', async () => {
		const el = await mount({ 'workspace-id': 'ta-shell', 'user-name': 'Ada', 'user-id': 'ada', 'server-config': LOCAL });
		const root = el.shadowRoot!;
		for (const tag of ['office-ui-app-rail', 'office-ui-channel-list', 'office-ui-chat-list', 'office-ui-chat-composer'])
			expect(root.querySelector(tag), tag).not.toBeNull();
		expect(root.querySelector('.topbar')!.textContent).toContain('Teams');
		expect(el.client).not.toBeNull();
		el.remove();
	});

	it('creates a channel, posts through the composer and shows the message', async () => {
		const el = await mount({ 'workspace-id': 'ta-post', 'user-name': 'Ada', 'user-id': 'ada', 'server-config': LOCAL });
		el.client!.createChannel('Ops');
		await tick(10);
		await el.updateComplete;
		const composer = q<HTMLElement & { updateComplete: Promise<boolean> }>(el.shadowRoot, 'office-ui-chat-composer');
		const area = q<HTMLTextAreaElement>(composer.shadowRoot, 'textarea');
		area.value = 'Deploy at <b>noon</b>';
		area.dispatchEvent(new Event('input'));
		await composer.updateComplete;
		area.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter' }));
		await tick(20);
		await el.updateComplete;
		const list = q<HTMLElement & { updateComplete: Promise<boolean> }>(el.shadowRoot, 'office-ui-chat-list');
		await list.updateComplete;
		expect(q(list.shadowRoot, '.text').textContent).toBe('Deploy at noon');
		expect(q(el.shadowRoot, '.channel-head h1').textContent).toContain('Ops');
		el.remove();
	});

	it('lets the host take over opening a file', async () => {
		const el = await mount({ 'workspace-id': 'ta-open', 'user-name': 'Ada', 'user-id': 'ada', 'server-config': LOCAL });
		const seen: unknown[] = [];
		el.addEventListener('teams-open-file', (e) => {
			seen.push((e as CustomEvent).detail.attachment.name);
			e.preventDefault();
		});
		el.client!.createChannel('Files');
		await tick(10);
		await el.client!.send({ text: 'see file', files: [new File(['x'], 'Plan.docx')] });
		await tick(20);
		await el.updateComplete;
		const list = q<HTMLElement & { updateComplete: Promise<boolean> }>(el.shadowRoot, 'office-ui-chat-list');
		await list.updateComplete;
		q<HTMLElement>(list.shadowRoot, '.file').click();
		await tick(5); // the link is requested before the event
		expect(seen).toEqual(['Plan.docx']);
		el.remove();
	});

	it('restarts the client when the workspace changes', async () => {
		const el = await mount({ 'workspace-id': 'ta-one', 'user-name': 'Ada', 'user-id': 'ada', 'server-config': LOCAL });
		const first = el.client;
		el.setAttribute('workspace-id', 'ta-two');
		await el.updateComplete;
		expect(el.client).not.toBe(first);
		expect(el.client!.workspace.session.roomId).toBe('ta-two');
		el.remove();
	});
});

describe('bindings contract', () => {
	it('applies only changed props and reads the latest handlers', async () => {
		const el = await mount({ 'workspace-id': 'ta-bind', 'user-name': 'Ada', 'user-id': 'ada', 'server-config': LOCAL });
		const uploader = async () => ({ url: 'http://x/y' });
		applyTeamsProps(el, { workspaceId: 'ta-bind-2', uploadFile: uploader }, { workspaceId: 'ta-bind' });
		expect(el.workspaceId).toBe('ta-bind-2');
		expect(el.uploadFile).toBe(uploader);
		applyTeamsProps(el, {}, { uploadFile: uploader });
		expect(el.uploadFile).toBeUndefined();

		let handlers: { onReady?: (d: unknown) => void } = { onReady: () => {} };
		const calls: unknown[] = [];
		const stop = listenTeamsEvents(el, () => handlers);
		handlers = { onReady: (d) => calls.push(d) };
		el.dispatchEvent(new CustomEvent('teams-ready', { detail: { user: { id: 'u', name: 'U' } } }));
		expect(calls).toEqual([{ user: { id: 'u', name: 'U' } }]);
		stop();
		el.dispatchEvent(new CustomEvent('teams-ready', { detail: {} }));
		expect(calls).toHaveLength(1);
		el.remove();
	});
});

describe('settings helpers', () => {
	it('round-trips ICE servers through one-per-line text', () => {
		const text = 'stun:s.example:3478\nturn:t.example:3478,turns:t.example:5349 user pass';
		const parsed = parseIceLines(text) as { urls: string[]; username?: string }[];
		expect(parsed[1]).toEqual({ urls: ['turn:t.example:3478', 'turns:t.example:5349'], username: 'user', credential: 'pass' });
		expect(iceToText({ mode: 'local', iceServers: [{ urls: 'stun:s.example:3478' }, { urls: ['turn:t.example:3478'], username: 'u', credential: 'p' }] })).toBe(
			'stun:s.example:3478\nturn:t.example:3478 u p',
		);
	});
});
