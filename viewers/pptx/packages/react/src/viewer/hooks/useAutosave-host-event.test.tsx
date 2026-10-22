// @vitest-environment happy-dom
/**
 * `onAutosave` (issue 151, item 5): the host hears of each recovery snapshot with the bytes that
 * were written, the same payload Vue and Svelte emit.
 */
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

const { saveMock } = vi.hoisted(() => ({ saveMock: vi.fn() }));

vi.mock(import('ooxml-ui/pptx'), async (importOriginal) => ({
	...(await importOriginal()),
	saveAutosaveSnapshot: saveMock,
}));

const { useAutosave } = await import('./useAutosave');

let container: HTMLDivElement;
let root: Root;

beforeEach(() => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	vi.useFakeTimers();
	saveMock.mockReset();
	saveMock.mockResolvedValue(true);
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => {
		root.unmount();
	});
	container.remove();
	vi.useRealTimers();
	globalThis.IS_REACT_ACT_ENVIRONMENT = false;
});

describe('useAutosave onSaved', () => {
	it('passes the written snapshot bytes to the host', async () => {
		const bytes = new Uint8Array([9, 8, 7]);
		const onSaved = vi.fn();
		function Probe(): null {
			useAutosave({
				isDirty: true,
				filePath: 'deck.pptx',
				serializeSlides: async () => bytes,
				intervalMs: 1000,
				onSaved,
			});
			return null;
		}
		act(() => root.render(<Probe />));
		await act(async () => {
			await vi.advanceTimersByTimeAsync(1100);
		});
		expect(saveMock).toHaveBeenCalledWith('deck.pptx', bytes);
		expect(onSaved).toHaveBeenCalledExactlyOnceWith(bytes);
	});
});
