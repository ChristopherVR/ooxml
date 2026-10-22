import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import type { CollaborationConfig } from '../types';
import type { ExternalCollaborationSession } from './collaboration-external-session';

const mockDoc = {} as never;
const loadMock = vi.fn().mockResolvedValue({});
const saveMock = vi.fn().mockResolvedValue(new Uint8Array([0x50, 0x4b, 0x03, 0x04]));

vi.mock(import('ooxml-core/pptx'), async (importOriginal) => {
	const actual = await importOriginal();
	return {
		...actual,
		PptxHandler: vi.fn().mockImplementation(function PptxHandlerMock(this: object) {
			Object.assign(this, { load: loadMock, save: saveMock });
		}),
	};
});

vi.mock(import('./collaboration-sync'), async (importOriginal) => {
	const actual = await importOriginal();
	return {
		...actual,
		readSlidesFromYDoc: vi.fn().mockReturnValue([{ id: 's1', elements: [] }]),
	};
});

const { createWriteBackScheduler } = await import('./collaboration-writeback');

describe('createWriteBackScheduler getSaveOptions wiring', () => {
	beforeEach(() => {
		vi.clearAllMocks();
		loadMock.mockResolvedValue({});
		saveMock.mockResolvedValue(new Uint8Array([0x50, 0x4b, 0x03, 0x04]));
		vi.useFakeTimers();
	});
	afterEach(() => vi.useRealTimers());

	function config(overrides: Partial<CollaborationConfig> = {}): CollaborationConfig {
		return {
			role: 'owner',
			onWriteBack: vi.fn(),
			writeBackDebounceMs: 0,
			...overrides,
		} as unknown as CollaborationConfig;
	}

	function externalConfig(initiallySynced = true) {
		let synced = initiallySynced;
		const listeners = new Set<() => void>();
		const externalSession = {
			getSnapshot: () => ({ status: 'disconnected', synced }),
			subscribe: (listener: () => void) => {
				listeners.add(listener);
				return () => listeners.delete(listener);
			},
		} as ExternalCollaborationSession;
		return {
			config: config({ externalSession }),
			listeners,
			setSynced: (value: boolean) => {
				synced = value;
				for (const listener of listeners) {
					listener();
				}
			},
		};
	}

	it('only schedules host snapshots once synced, including synced offline sessions', async () => {
		const host = externalConfig(false);
		const scheduler = createWriteBackScheduler({
			getYDoc: () => mockDoc,
			getSourceBytes: () => new Uint8Array([1]),
		});
		scheduler.schedule(host.config);
		await vi.runAllTimersAsync();
		expect(loadMock).not.toHaveBeenCalled();
		host.setSynced(true);
		scheduler.schedule(host.config);
		await vi.runAllTimersAsync();
		expect(host.config.onWriteBack).toHaveBeenCalledOnce();
		expect(host.listeners.size).toBe(0);
	});

	it('uses a retained serializer without reloading source bytes', async () => {
		const bytes = new Uint8Array([7]);
		const serialize = vi.fn((isCurrent: () => boolean) => {
			expect(isCurrent()).toBeTruthy();
			return bytes;
		});
		const scheduler = createWriteBackScheduler({ getYDoc: () => mockDoc, serialize });
		const session = config();
		scheduler.schedule(session);
		await vi.runAllTimersAsync();
		expect(loadMock).not.toHaveBeenCalled();
		expect(serialize).toHaveBeenCalledOnce();
		expect(session.onWriteBack).toHaveBeenCalledWith(bytes);
	});

	it.each(['source', 'retained'] as const)(
		'contains an async host callback rejection with the %s serializer',
		async (mode) => {
			const bytes = new Uint8Array([7]);
			let attempts = 0;
			const onWriteBack = async (): Promise<void> => {
				attempts++;
				if (attempts === 1) {
					throw new Error('Host save failed');
				}
			};
			const scheduler = createWriteBackScheduler({
				getYDoc: () => mockDoc,
				getSourceBytes: () => bytes,
				...(mode === 'retained' ? { serialize: () => bytes } : {}),
			});
			const session = config({ onWriteBack });
			scheduler.schedule(session);
			await vi.runAllTimersAsync();
			scheduler.schedule(session);
			await vi.runAllTimersAsync();
			expect(attempts).toBe(2);
		},
	);

	it.each(['cancel', 'replace', 'unsync'] as const)(
		'revokes a retained serializer after %s',
		async (cause) => {
			let finish!: (bytes: Uint8Array) => void;
			let isCurrent!: () => boolean;
			let document = mockDoc;
			const host = externalConfig();
			const scheduler = createWriteBackScheduler({
				getYDoc: () => document,
				serialize: (check) => {
					isCurrent = check;
					return new Promise((resolve) => {
						finish = resolve;
					});
				},
			});
			scheduler.schedule(host.config);
			vi.advanceTimersByTime(0);
			expect(isCurrent()).toBeTruthy();
			if (cause === 'cancel') {
				scheduler.cancel();
			}
			if (cause === 'replace') {
				document = {} as never;
			}
			if (cause === 'unsync') {
				host.setSynced(false);
				host.setSynced(true);
			}
			expect(isCurrent()).toBeFalsy();
			finish(new Uint8Array([7]));
			await vi.advanceTimersByTimeAsync(0);
			expect(host.config.onWriteBack).not.toHaveBeenCalled();
			expect(host.listeners.size).toBe(0);
		},
	);

	it('cancels a pending host snapshot when sync is lost', async () => {
		const host = externalConfig();
		const scheduler = createWriteBackScheduler({
			getYDoc: () => mockDoc,
			getSourceBytes: () => new Uint8Array([1]),
		});
		scheduler.schedule(host.config);
		host.setSynced(false);
		host.setSynced(true);
		await vi.runAllTimersAsync();
		expect(loadMock).not.toHaveBeenCalled();
		expect(host.listeners.size).toBe(0);
	});

	it.each(['load', 'save'])(
		'invalidates an in-flight %s even if the host re-syncs before it finishes',
		async (stage) => {
			let finish!: () => void;
			const deferred = new Promise((resolve) => {
				finish = () => resolve(new Uint8Array([2]));
			});
			(stage === 'load' ? loadMock : saveMock).mockReturnValue(deferred);
			const host = externalConfig();
			const scheduler = createWriteBackScheduler({
				getYDoc: () => mockDoc,
				getSourceBytes: () => new Uint8Array([1]),
			});
			scheduler.schedule(host.config);
			vi.advanceTimersByTime(0);
			await Promise.resolve();
			host.setSynced(false);
			host.setSynced(true);
			finish();
			await Promise.resolve();
			await Promise.resolve();
			expect(host.config.onWriteBack).not.toHaveBeenCalled();
			if (stage === 'load') {
				expect(saveMock).not.toHaveBeenCalled();
			}
			expect(host.listeners.size).toBe(0);
		},
	);

	it('passes getSaveOptions() through to handler.save so session-level edits survive write-back', async () => {
		const saveOptions = { viewProperties: { showComments: true }, tags: [{ tags: [] }] };
		const scheduler = createWriteBackScheduler({
			getYDoc: () => mockDoc,
			getSourceBytes: () => new Uint8Array([1, 2, 3]),
			getSaveOptions: () => saveOptions as never,
		});

		scheduler.schedule(config());
		await vi.runAllTimersAsync();

		expect(saveMock).toHaveBeenCalledExactlyOnceWith([{ id: 's1', elements: [] }], saveOptions);
		vi.useRealTimers();
	});

	it('calls handler.save with undefined options when getSaveOptions is not supplied (pre-existing behaviour)', async () => {
		const scheduler = createWriteBackScheduler({
			getYDoc: () => mockDoc,
			getSourceBytes: () => new Uint8Array([1, 2, 3]),
		});

		scheduler.schedule(config());
		await vi.runAllTimersAsync();

		expect(saveMock).toHaveBeenCalledExactlyOnceWith([{ id: 's1', elements: [] }], undefined);
		vi.useRealTimers();
	});

	it('does not finish an in-flight write-back after its session was stopped', async () => {
		let finishSave!: (bytes: Uint8Array) => void;
		saveMock.mockReturnValue(
			new Promise<Uint8Array>((resolve) => {
				finishSave = resolve;
			}),
		);
		const scheduler = createWriteBackScheduler({
			getYDoc: () => mockDoc,
			getSourceBytes: () => new Uint8Array([1]),
		});
		const session = config();
		scheduler.schedule(session);
		vi.advanceTimersByTime(0);
		await Promise.resolve();
		expect(saveMock).toHaveBeenCalledOnce();
		scheduler.cancel();
		finishSave(new Uint8Array([2]));
		await Promise.resolve();
		expect(session.onWriteBack).not.toHaveBeenCalled();
	});

	it('does not read replacement-session options after an obsolete source load completes', async () => {
		let finishLoad!: (value: object) => void;
		loadMock.mockReturnValue(
			new Promise<object>((resolve) => {
				finishLoad = resolve;
			}),
		);
		const getSaveOptions = vi.fn();
		const scheduler = createWriteBackScheduler({
			getYDoc: () => mockDoc,
			getSourceBytes: () => new Uint8Array([1]),
			getSaveOptions,
		});
		const session = config();
		scheduler.schedule(session);
		vi.advanceTimersByTime(0);
		scheduler.cancel();
		finishLoad({});
		await Promise.resolve();
		expect(saveMock).not.toHaveBeenCalled();
		expect(getSaveOptions).not.toHaveBeenCalled();
		expect(session.onWriteBack).not.toHaveBeenCalled();
	});
});

describe('createWriteBackScheduler acknowledgement, retry, status and page close', () => {
	beforeEach(() => {
		vi.clearAllMocks();
		vi.useFakeTimers();
	});
	afterEach(() => vi.useRealTimers());

	const owner = (overrides: Partial<CollaborationConfig> = {}) =>
		({ role: 'owner', writeBackDebounceMs: 10, ...overrides }) as unknown as CollaborationConfig;

	it('treats a returned promise as the acknowledgement and reports the status', async () => {
		let release!: () => void;
		const onWriteBack = vi.fn(
			() =>
				new Promise<void>((resolve) => {
					release = resolve;
				}),
		);
		const onWriteBackStatus = vi.fn();
		const scheduler = createWriteBackScheduler({
			getYDoc: () => mockDoc,
			serialize: () => new Uint8Array([1]),
			onWriteBackStatus,
		});
		expect(scheduler.getState()).toBeUndefined();
		scheduler.schedule(owner({ onWriteBack }));
		expect(scheduler.getState()?.status).toBe('pending');
		await vi.advanceTimersByTimeAsync(10);
		expect(scheduler.getState()?.status).toBe('saving');
		release();
		await vi.advanceTimersByTimeAsync(0);
		expect(scheduler.getState()?.status).toBe('saved');
		expect(scheduler.getState()?.lastSavedAt).toBeTypeOf('number');
		expect(onWriteBackStatus).toHaveBeenCalled();
		scheduler.cancel();
	});

	it('retries a rejection per writeBackRetry, then reports failed with the message', async () => {
		const onWriteBack = vi.fn().mockRejectedValue(new Error('offline'));
		const scheduler = createWriteBackScheduler({
			getYDoc: () => mockDoc,
			serialize: () => new Uint8Array([1]),
		});
		scheduler.schedule(owner({ onWriteBack, writeBackRetry: { attempts: 3, baseDelayMs: 20 } }));
		await vi.advanceTimersByTimeAsync(1_000);
		expect(onWriteBack).toHaveBeenCalledTimes(3);
		expect(scheduler.getState()).toEqual({ status: 'failed', error: 'offline' });
		scheduler.cancel();
	});

	it('a serializer that stops because the session changed is not a failed save', async () => {
		const onWriteBack = vi.fn();
		const scheduler = createWriteBackScheduler({
			getYDoc: () => mockDoc,
			serialize: (isCurrent) => {
				scheduler.cancel();
				expect(isCurrent()).toBe(false);
				throw new Error('Collaboration session changed during serialization');
			},
		});
		scheduler.schedule(owner({ onWriteBack, writeBackDebounceMs: 0 }));
		await vi.advanceTimersByTimeAsync(10);
		expect(onWriteBack).not.toHaveBeenCalled();
		expect(scheduler.getState()?.status).not.toBe('failed');
	});

	it('flushes on pagehide and when hidden, and stops listening after cancel', async () => {
		const listeners = new Map<string, () => void>();
		const doc = {
			visibilityState: 'visible',
			addEventListener: (type: string, listener: () => void) => void listeners.set(type, listener),
			removeEventListener: (type: string) => void listeners.delete(type),
		};
		vi.stubGlobal('addEventListener', (type: string, listener: () => void) =>
			listeners.set(type, listener),
		);
		vi.stubGlobal('removeEventListener', (type: string) => listeners.delete(type));
		vi.stubGlobal('document', doc);
		try {
			const onWriteBack = vi.fn();
			const scheduler = createWriteBackScheduler({
				getYDoc: () => mockDoc,
				serialize: () => new Uint8Array([1]),
			});
			scheduler.schedule(owner({ onWriteBack, writeBackDebounceMs: 60_000 }));
			listeners.get('pagehide')!();
			await vi.advanceTimersByTimeAsync(0);
			expect(onWriteBack).toHaveBeenCalledTimes(1);

			scheduler.schedule(owner({ onWriteBack, writeBackDebounceMs: 60_000 }));
			doc.visibilityState = 'hidden';
			listeners.get('visibilitychange')!();
			await vi.advanceTimersByTimeAsync(0);
			expect(onWriteBack).toHaveBeenCalledTimes(2);

			scheduler.schedule(owner({ onWriteBack, writeBackDebounceMs: 60_000 }));
			scheduler.cancel();
			expect(listeners.size).toBe(0);
			expect(onWriteBack).toHaveBeenCalledTimes(2);
		} finally {
			vi.unstubAllGlobals();
		}
	});

	it('flush() resolves after the host acknowledged', async () => {
		const onWriteBack = vi.fn(() => Promise.resolve());
		const scheduler = createWriteBackScheduler({
			getYDoc: () => mockDoc,
			serialize: () => new Uint8Array([1]),
		});
		scheduler.schedule(owner({ onWriteBack, writeBackDebounceMs: 60_000 }));
		await scheduler.flush();
		expect(onWriteBack).toHaveBeenCalledTimes(1);
		expect(scheduler.getState()?.status).toBe('saved');
		scheduler.cancel();
	});
});

describe('createWriteBackScheduler flush on stop', () => {
	beforeEach(() => {
		vi.clearAllMocks();
		vi.useFakeTimers();
	});
	afterEach(() => vi.useRealTimers());

	it('writes an unsaved owner change once when cancelled, with the document as it was', async () => {
		let doc: typeof mockDoc | null = mockDoc;
		const onWriteBack = vi.fn();
		const serialize = vi.fn(() => new Uint8Array([9]));
		const scheduler = createWriteBackScheduler({ getYDoc: () => doc, serialize });
		scheduler.schedule({ role: 'owner', onWriteBack, writeBackDebounceMs: 60_000 });
		scheduler.cancel();
		doc = null; // a binding clears its document right after cancelling
		await vi.advanceTimersByTimeAsync(0);
		expect(onWriteBack).toHaveBeenCalledExactlyOnceWith(new Uint8Array([9]));
		await vi.advanceTimersByTimeAsync(120_000);
		expect(onWriteBack).toHaveBeenCalledOnce();
	});

	it('writes nothing when nothing is unsaved, for a non-owner, or after readiness was lost', async () => {
		const onWriteBack = vi.fn();
		const serialize = () => new Uint8Array([9]);
		const idle = createWriteBackScheduler({ getYDoc: () => mockDoc, serialize });
		idle.cancel();
		const saved = createWriteBackScheduler({ getYDoc: () => mockDoc, serialize });
		saved.schedule({ role: 'owner', onWriteBack, writeBackDebounceMs: 0 });
		await vi.advanceTimersByTimeAsync(10);
		expect(onWriteBack).toHaveBeenCalledOnce();
		saved.cancel();
		await vi.advanceTimersByTimeAsync(10);
		expect(onWriteBack).toHaveBeenCalledOnce();

		const viewer = createWriteBackScheduler({ getYDoc: () => mockDoc, serialize });
		viewer.schedule({ role: 'collaborator', onWriteBack, writeBackDebounceMs: 60_000 });
		viewer.cancel();

		const host = (() => {
			let synced = true;
			const session = {
				getSnapshot: () => ({ status: 'connected', synced }),
				subscribe: () => () => {},
			} as unknown as ExternalCollaborationSession;
			return { session, lose: () => (synced = false) };
		})();
		const suspended = createWriteBackScheduler({ getYDoc: () => mockDoc, serialize });
		suspended.schedule({
			role: 'owner',
			onWriteBack,
			writeBackDebounceMs: 60_000,
			externalSession: host.session,
		});
		host.lose();
		suspended.cancel();
		await vi.advanceTimersByTimeAsync(10);
		expect(onWriteBack).toHaveBeenCalledOnce();
	});
});
