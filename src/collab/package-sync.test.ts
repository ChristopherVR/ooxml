import { afterEach, describe, expect, it } from 'vitest';
import { bindDocument } from './binding.js';
import {
	createBroadcastTransport,
	type BroadcastChannelConstructor,
	type BroadcastChannelLike,
} from './broadcast-transport.js';
import { packageAdapter, type SharedPackage } from './package-adapter.js';
import { createCollabSession, type CollabSession } from './session.js';
import { transportProvider } from './transport-provider.js';

/** An in-process BroadcastChannel: delivers to every other channel of the same name, async. */
function fakeChannels(): BroadcastChannelConstructor {
	const open = new Map<string, Set<BroadcastChannelLike>>();
	return class implements BroadcastChannelLike {
		onmessage: ((event: { data: unknown }) => void) | null = null;
		constructor(private readonly name: string) {
			if (!open.has(name)) open.set(name, new Set());
			open.get(name)!.add(this);
		}
		postMessage(message: unknown): void {
			const data = structuredClone(message);
			for (const other of open.get(this.name) ?? [])
				if (other !== this) queueMicrotask(() => other.onmessage?.({ data }));
		}
		close(): void {
			open.get(this.name)?.delete(this);
		}
	};
}

const settle = () => new Promise((resolve) => setTimeout(resolve, 10));
const sessions: CollabSession[] = [];
afterEach(() => sessions.splice(0).forEach((session) => session.destroy()));

function join(Channel: BroadcastChannelConstructor, name: string) {
	const session = createCollabSession({
		roomId: 'drawing',
		provider: transportProvider({
			transport: createBroadcastTransport({ roomId: 'drawing', BroadcastChannel: Channel }),
		}),
		user: { name },
		heartbeatMs: 0,
		syncGraceMs: 0,
	});
	sessions.push(session);
	return session;
}

describe('same-browser package sharing', () => {
	it('seeds, adopts and republishes whole packages between tabs without echoes', async () => {
		const Channel = fakeChannels();
		const adapter = packageAdapter();
		const a = join(Channel, 'Ada');
		const seen: { a: SharedPackage[]; b: SharedPackage[] } = { a: [], b: [] };
		const bindA = bindDocument(a, adapter, {
			getLocalModel: () => ({ bytes: new Uint8Array([1, 2]), revision: 0 }),
			onRemoteModel: (model) => seen.a.push(model),
		});
		await settle();
		const b = join(Channel, 'Grace');
		const bindB = bindDocument(b, adapter, {
			getLocalModel: () => ({ bytes: new Uint8Array([9]), revision: 0 }),
			onRemoteModel: (model) => seen.b.push(model),
		});
		await settle();
		expect(seen.b.at(-1)).toEqual({ bytes: new Uint8Array([1, 2]), revision: 1 });
		expect(b.peers().map((peer) => peer.userName)).toEqual(['Ada']);

		expect(bindB.push({ bytes: new Uint8Array([3]), revision: 1 })).toBe(true);
		await settle();
		expect(seen.a.at(-1)).toEqual({ bytes: new Uint8Array([3]), revision: 2 });
		// B never adopts its own push.
		expect(seen.b.some((model) => model.bytes[0] === 3)).toBe(false);
		bindA.dispose();
		bindB.dispose();
	});

	it('refuses packages above the limit', () => {
		const Channel = fakeChannels();
		const session = join(Channel, 'Ada');
		const adapter = packageAdapter({ maxBytes: 2 });
		expect(() =>
			adapter.write(session.doc, { bytes: new Uint8Array(3), revision: 0 }, Symbol() as never),
		).toThrow(/exceeds/);
		expect(adapter.isEmpty(session.doc)).toBe(true);
	});

	it('reports a missing BroadcastChannel as an error instead of throwing', () => {
		const transport = createBroadcastTransport({
			roomId: 'drawing',
			BroadcastChannel: undefined as never,
		});
		const errors: Error[] = [];
		const scope = globalThis as { BroadcastChannel?: unknown };
		const saved = scope.BroadcastChannel;
		delete scope.BroadcastChannel;
		try {
			createBroadcastTransport({ roomId: 'drawing' }).connect({
				open: () => {},
				close: () => {},
				message: () => {},
				error: (error) => errors.push(error),
			});
		} finally {
			if (saved) scope.BroadcastChannel = saved;
		}
		expect(errors[0]?.message).toMatch(/BroadcastChannel/);
		expect(() => transport.send(new Uint8Array([1]))).not.toThrow();
	});
});
