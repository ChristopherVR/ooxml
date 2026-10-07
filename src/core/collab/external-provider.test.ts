import { describe, expect, it, vi } from 'vitest';
import { adaptYjsProvider, type YjsProviderLike } from './external-provider';

function mockProvider(initial: Partial<YjsProviderLike> = {}) {
	const listeners = new Map<string, Set<(payload: never) => void>>();
	const provider: YjsProviderLike = {
		on: (event, listener) => {
			let entries = listeners.get(event);
			if (!entries) listeners.set(event, (entries = new Set()));
			entries.add(listener);
		},
		off: (event, listener) => {
			listeners.get(event)?.delete(listener);
		},
		connect: vi.fn(),
		disconnect: vi.fn(),
		destroy: vi.fn(),
		...initial,
	};
	return {
		provider,
		emit: (event: string, payload: unknown) => {
			for (const listener of listeners.get(event) ?? []) listener(payload as never);
		},
		listeners,
	};
}

describe('external Yjs provider lifecycle', () => {
	it('interprets both boolean and object sync events without opening on false', () => {
		const source = mockProvider();
		const adapter = adaptYjsProvider(source.provider);
		for (const event of ['sync', 'synced']) {
			source.emit(event, true);
			expect(adapter.synced).toBe(true);
			source.emit(event, { synced: false });
			expect(adapter.synced).toBe(false);
			source.emit(event, { synced: true });
			expect(adapter.synced).toBe(true);
			source.emit(event, false);
			expect(adapter.synced).toBe(false);
		}
		adapter.destroy();
	});

	it('resets sync on WebRTC disconnect and exposes explicit disconnect immediately', () => {
		const source = mockProvider({ connected: true, synced: true });
		const adapter = adaptYjsProvider(source.provider);
		source.emit('status', { connected: false });
		expect(adapter.status).toBe('disconnected');
		expect(adapter.synced).toBe(false);
		source.emit('status', { connected: true });
		source.emit('synced', { synced: true });
		adapter.disconnect();
		expect(adapter.status).toBe('disconnected');
		expect(adapter.synced).toBe(false);
		expect(source.provider.disconnect).toHaveBeenCalledOnce();
		adapter.destroy();
	});

	it('recognises a connecting socket and destroys once without later reconnects', () => {
		const source = mockProvider({ wsconnecting: true });
		const adapter = adaptYjsProvider(source.provider);
		expect(adapter.status).toBe('connecting');
		adapter.connect();
		expect(source.provider.connect).not.toHaveBeenCalled();
		adapter.destroy();
		adapter.destroy();
		adapter.connect();
		adapter.disconnect();
		expect(source.provider.destroy).toHaveBeenCalledOnce();
		expect([...source.listeners.values()].every((entries) => entries.size === 0)).toBe(true);
	});
});
