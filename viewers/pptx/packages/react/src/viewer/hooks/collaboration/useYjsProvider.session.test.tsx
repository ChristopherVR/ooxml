// @vitest-environment happy-dom
/**
 * useYjsProvider drives the shared collaboration session controller: one
 * room lifecycle for every binding. These render the real hook against a fake
 * y-webrtc provider.
 */
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const { created } = vi.hoisted(() => ({
	created: [] as Array<{ destroy: ReturnType<typeof vi.fn>; args: unknown[] }>,
}));

vi.mock(import('y-webrtc'), () => ({
	WebrtcProvider: class {
		awareness = {
			clientID: 42,
			getStates: () => new Map(),
			getLocalState: () => null,
			setLocalState: () => {},
			setLocalStateField: () => {},
			on: () => {},
			off: () => {},
		};
		destroy = vi.fn();
		constructor(...args: unknown[]) {
			created.push({ destroy: this.destroy, args });
		}
		on() {}
	} as never,
}));

const { useYjsProvider } = await import('./useYjsProvider');
type Result = ReturnType<typeof useYjsProvider>;

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	created.length = 0;
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
	globalThis.IS_REACT_ACT_ENVIRONMENT = false;
});

describe('useYjsProvider over the shared session controller', () => {
	it('opens a webrtc room, reports it connected and leaves it on unmount', async () => {
		let latest: Result | undefined;
		function Probe(): null {
			latest = useYjsProvider({
				config: { roomId: 'room-1', serverUrl: '', transport: 'webrtc', userName: 'Ada' },
			});
			return null;
		}
		await act(async () => {
			root.render(<Probe />);
		});
		await act(async () => {
			await vi.dynamicImportSettled();
		});
		expect(created).toHaveLength(1);
		expect(latest?.status).toBe('connected');
		expect(latest?.clientId).toBe(42);
		expect(latest?.doc).not.toBeNull();
		act(() => root.unmount());
		expect(created[0]?.destroy).toHaveBeenCalledOnce();
		root = createRoot(container);
	});

	it('stays dormant without a config and reports an invalid room as an error', async () => {
		let latest: Result | undefined;
		function Probe({ roomId }: { roomId?: string }): null {
			latest = useYjsProvider({
				config:
					roomId === undefined
						? undefined
						: { roomId, serverUrl: '', transport: 'webrtc', userName: 'Ada' },
			});
			return null;
		}
		await act(async () => {
			root.render(<Probe />);
		});
		expect(latest?.status).toBe('disconnected');
		await act(async () => {
			root.render(<Probe roomId='' />);
		});
		expect(latest?.status).toBe('error');
		expect(created).toHaveLength(0);
	});
});
