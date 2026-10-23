/**
 * collaboration.service.test.ts: Unit tests for the CollaborationService
 * connect() reentrancy guard.
 *
 * A guest can trigger a second connect() while the first is still awaiting the
 * dynamic transport import (e.g. the host config effect re-running when the
 * placeholder deck loads). Without the token guard, the second provider join on
 * the same room throws inside Yjs and the catch handler tears down the
 * surviving session. These tests mock the transport factories with manually
 * resolved promises so the interleaving is deterministic.
 *
 * The service is constructed via `Injector.create` with a stubbed `DestroyRef`
 * (same minimal-injection-context pattern as `print.service.test.ts`).
 */

import { readFile } from 'node:fs/promises';
import { resolve as resolvePath } from 'node:path';

import { DestroyRef, Injector } from '@angular/core';
import type { CollabYSession, CollaborationConfig, YjsFactories } from 'ooxml-ui/pptx';
import { PptxHandler } from 'pptx-viewer-core';
import { beforeEach, describe, expect, it, vi } from 'vitest';
import * as Y from 'yjs';

import type { AwarenessLike } from './collaboration-providers';
import { CollaborationService } from './collaboration.service';

const mocks = vi.hoisted(() => {
	const pending: Array<{
		resolve: (session: CollabYSession) => void;
		reject: (error: unknown) => void;
	}> = [];
	return { pending };
});

vi.mock(import('./collaboration-providers'), async (importOriginal) => {
	const actual = await importOriginal();
	// Owned transports resolve by hand; a borrowed host session opens for real.
	const open: typeof actual.openCollaborationSession = (config, transport) =>
		config.externalSession
			? actual.openCollaborationSession(config, transport)
			: new Promise((resolve, reject) => {
					mocks.pending.push({ resolve, reject });
				});
	return { ...actual, openCollaborationSession: open };
});

interface FakeBundle {
	bundle: CollabYSession;
	dispose: ReturnType<typeof vi.fn>;
}

function makeBundle(): FakeBundle {
	const slidesArray = { observeDeep: vi.fn(), unobserveDeep: vi.fn(), length: 0 };
	const doc = { getArray: () => slidesArray, destroy: vi.fn() };
	const awareness: AwarenessLike = {
		clientID: 1,
		setLocalStateField: vi.fn(),
		getStates: () => new Map<number, Record<string, unknown>>(),
		on: vi.fn(),
		off: vi.fn(),
	};
	const provider = {
		awareness,
		onStatus: vi.fn(),
		onSynced: vi.fn(),
		connectedNow: true,
		syncedNow: false,
		destroy: vi.fn(),
	};
	const factories = {
		createMap: vi.fn(),
		createArray: vi.fn(),
		createText: vi.fn(),
	} as unknown as YjsFactories;
	const dispose = vi.fn();
	return {
		bundle: {
			doc: doc as unknown as CollabYSession['doc'],
			provider,
			awareness,
			factories,
			dispose,
		},
		dispose,
	};
}

function createService(): CollaborationService {
	const injector = Injector.create({
		providers: [
			{ provide: DestroyRef, useValue: { onDestroy: () => () => undefined } as DestroyRef },
			{ provide: CollaborationService, useClass: CollaborationService },
		],
	});
	return injector.get(CollaborationService);
}

function config(roomId: string): CollaborationConfig {
	return { roomId, serverUrl: '', transport: 'webrtc', userName: 'Tester' };
}

describe('collaborationService connect reentrancy', () => {
	beforeEach(() => {
		mocks.pending.length = 0;
	});

	it('writes back an already-synced owner startup deck without a later edit', async () => {
		const source = new Uint8Array(
			await readFile(
				resolvePath(__dirname, '../../../../../../e2e/pptx/fixtures/text-layout.pptx'),
			),
		);
		const data = await new PptxHandler().load(source.buffer);
		const doc = new Y.Doc();
		const svc = createService();
		const onWriteBack = vi.fn();
		try {
			const pending = svc.connect(
				{
					...config('owner-startup'),
					role: 'owner',
					sessionIntent: 'create',
					writeBackDebounceMs: 0,
					onWriteBack,
					externalSession: {
						doc,
						getSnapshot: () => ({ status: 'connected', synced: true }),
						subscribe: () => () => {},
						awareness: {
							clientID: doc.clientID,
							getLocalState: () => ({}),
							setLocalState: vi.fn(),
							setLocalStateField: vi.fn(),
							getStates: () => new Map(),
							on: vi.fn(),
							off: vi.fn(),
						},
					},
				},
				{ getSourceBytes: () => source },
			);
			svc.seedBaseline(data.slides);
			await pending;
			await vi.waitFor(() => expect(onWriteBack).toHaveBeenCalledOnce(), { timeout: 10_000 });
			const bytes = onWriteBack.mock.calls[0][0] as Uint8Array;
			const reopened = await new PptxHandler().load(bytes.buffer as ArrayBuffer);
			expect(reopened.slides).toHaveLength(data.slides.length);
		} finally {
			svc.disconnect();
			doc.destroy();
		}
	});

	it('a superseding connect destroys the first bundle and keeps the second', async () => {
		const svc = createService();
		const first = svc.connect(config('room-1'));
		const second = svc.connect(config('room-1'));
		expect(mocks.pending).toHaveLength(2);

		const b1 = makeBundle();
		const b2 = makeBundle();
		mocks.pending[0].resolve(b1.bundle);
		mocks.pending[1].resolve(b2.bundle);
		await Promise.all([first, second]);

		// First (stale) bundle torn down without touching the newer session.
		expect(b1.dispose).toHaveBeenCalledOnce();
		expect(b2.dispose).not.toHaveBeenCalled();
		expect(svc.active()).toBeTruthy();
		expect(svc.status()).toBe('connected');
		svc.disconnect();
	});

	it('disconnect during a pending connect discards the created bundle', async () => {
		const svc = createService();
		const pendingConnect = svc.connect(config('room-2'));
		svc.disconnect();

		const b = makeBundle();
		mocks.pending[0].resolve(b.bundle);
		await pendingConnect;

		expect(b.dispose).toHaveBeenCalledOnce();
		expect(svc.active()).toBeFalsy();
		expect(svc.status()).toBe('disconnected');
	});

	it('a rejected superseded connect leaves the newer session alone', async () => {
		const svc = createService();
		const first = svc.connect(config('room-3'));
		const second = svc.connect(config('room-3'));

		const b2 = makeBundle();
		mocks.pending[0].reject(new Error('transport import failed'));
		mocks.pending[1].resolve(b2.bundle);
		await Promise.all([first, second]);

		expect(svc.status()).toBe('connected');
		expect(svc.active()).toBeTruthy();
		expect(b2.dispose).not.toHaveBeenCalled();
		svc.disconnect();
	});
});
