import type { AddressInfo } from 'node:net';

import { afterEach, describe, expect, it } from 'vitest';
import { WebSocket } from 'ws';
import { WebsocketProvider } from 'y-websocket';
import * as Y from 'yjs';

import { createReviewHub } from './app';

const hubs: ReturnType<typeof createReviewHub>[] = [];
const providers: WebsocketProvider[] = [];

afterEach(() => {
	for (const provider of providers.splice(0)) provider.destroy();
	for (const hub of hubs.splice(0)) hub.close();
});

async function startHub() {
	const hub = createReviewHub({ secret: 'test', source: new Uint8Array([1, 2, 3]), ai: null });
	hubs.push(hub);
	await new Promise<void>((resolve) => hub.server.listen(0, '127.0.0.1', resolve));
	const base = `http://127.0.0.1:${(hub.server.address() as AddressInfo).port}`;
	const as = async (account: string) => {
		const signIn = await fetch(`${base}/api/session`, {
			method: 'POST',
			headers: { 'x-review-hub': '1', 'Content-Type': 'application/json' },
			body: JSON.stringify({ account }),
		});
		const cookie = (signIn.headers.get('set-cookie') ?? '').split(';')[0] ?? '';
		const call = (path: string, init: RequestInit = {}) =>
			fetch(`${base}${path}`, {
				...init,
				headers: { cookie, 'x-review-hub': '1', ...init.headers },
			});
		return {
			call,
			grant: async () =>
				(await (await call('/api/reviews/q3-launch/ticket', { method: 'POST' })).json()) as {
					ticket: string;
					intent: string;
					role: string;
				},
		};
	};
	return { base, hub, as };
}

function connect(base: string, ticket: string) {
	const doc = new Y.Doc();
	const provider = new WebsocketProvider(base.replace('http', 'ws') + '/collab', 'q3-launch', doc, {
		WebSocketPolyfill: WebSocket as unknown as typeof globalThis.WebSocket,
		disableBc: true,
		params: { ticket },
	});
	providers.push(provider);
	const synced = new Promise<void>((resolve) => {
		if (provider.synced) resolve();
		provider.on('sync', (value: boolean) => value && resolve());
	});
	return { doc, provider, synced };
}

const until = async (check: () => boolean, ms = 3000): Promise<void> => {
	const end = Date.now() + ms;
	while (!check()) {
		if (Date.now() > end) throw new Error('timed out');
		await new Promise((resolve) => setTimeout(resolve, 20));
	}
};

describe('review hub', () => {
	it('hides reviews from strangers and refuses calls without the request header', async () => {
		const { base, as } = await startHub();
		expect((await fetch(`${base}/api/reviews/q3-launch`)).status).toBe(401);
		const ben = await as('ben');
		expect((await ben.call('/api/reviews/nope')).status).toBe(404);
		const record = (await (await ben.call('/api/reviews/q3-launch')).json()) as {
			permission: string;
			ai: boolean;
		};
		expect(record).toMatchObject({ permission: 'review', ai: false });
		const bare = await fetch(`${base}/api/reviews/q3-launch/ticket`, { method: 'POST' });
		expect(bare.status).toBe(403);
	});

	it('grants the seed to one writer and the read-only role to viewers', async () => {
		const { as } = await startHub();
		const [ana, ben, dee] = await Promise.all([as('ana'), as('ben'), as('dee')]);
		expect(await ana.grant()).toMatchObject({ role: 'owner', intent: 'create' });
		expect(await ben.grant()).toMatchObject({ role: 'collaborator', intent: 'join' });
		expect(await dee.grant()).toMatchObject({ role: 'viewer', intent: 'join' });
	});

	it("relays reviewers' edits to viewers and drops viewers' edits", async () => {
		const { base, as } = await startHub();
		const ben = connect(base, (await (await as('ben')).grant()).ticket);
		const dee = connect(base, (await (await as('dee')).grant()).ticket);
		await Promise.all([ben.synced, dee.synced]);
		ben.doc.getMap('deck').set('title', 'From Ben');
		await until(() => dee.doc.getMap('deck').get('title') === 'From Ben');
		dee.doc.getMap('deck').set('title', 'From Dee');
		// A late joiner sees the room as the relay holds it: Ben's value, not Dee's.
		const chen = connect(base, (await (await as('chen')).grant()).ticket);
		await chen.synced;
		await new Promise((resolve) => setTimeout(resolve, 100));
		expect(chen.doc.getMap('deck').get('title')).toBe('From Ben');
		expect(ben.doc.getMap('deck').get('title')).toBe('From Ben');
	});

	it('refuses sockets without a valid ticket for the room', async () => {
		const { base } = await startHub();
		const socket = new WebSocket(`${base.replace('http', 'ws')}/collab/q3-launch?ticket=forged`);
		const outcome = await new Promise<string>((resolve) => {
			socket.on('open', () => resolve('open'));
			socket.on('error', () => resolve('refused'));
		});
		expect(outcome).toBe('refused');
	});

	it("closes an account's sockets when the owner changes its access", async () => {
		const { base, as } = await startHub();
		const ana = await as('ana');
		const ben = connect(base, (await (await as('ben')).grant()).ticket);
		await ben.synced;
		const closed = new Promise<number | undefined>((resolve) =>
			ben.provider.on('connection-close', (event: CloseEvent | null) => resolve(event?.code)),
		);
		const change = await ana.call('/api/reviews/q3-launch/access/ben', {
			method: 'PUT',
			headers: { 'Content-Type': 'application/json' },
			body: JSON.stringify({ permission: 'view' }),
		});
		expect(change.status).toBe(200);
		expect(await closed).toBe(4403);
	});

	it('accepts versions from the owner only', async () => {
		const { as } = await startHub();
		const [ana, ben] = await Promise.all([as('ana'), as('ben')]);
		const upload = (who: typeof ana) =>
			who.call('/api/reviews/q3-launch/versions', { method: 'POST', body: new Uint8Array([7, 7]) });
		expect((await upload(ben)).status).toBe(403);
		expect((await upload(ana)).status).toBe(200);
		const saved = await ben.call('/api/reviews/q3-launch/versions/1');
		expect(new Uint8Array(await saved.arrayBuffer())).toEqual(new Uint8Array([7, 7]));
	});

	it('keeps the AI route closed when no model is configured', async () => {
		const { as } = await startHub();
		const ben = await as('ben');
		const chat = await ben.call('/api/ai/chat?review=q3-launch', {
			method: 'POST',
			body: '{"messages":[]}',
		});
		expect(chat.status).toBe(403);
	});
});
