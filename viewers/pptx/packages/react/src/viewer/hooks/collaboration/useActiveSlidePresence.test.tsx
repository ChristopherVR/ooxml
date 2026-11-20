// @vitest-environment happy-dom
/**
 * The local active slide reaches the session's presence from a layer that is not the editing
 * canvas, so a broadcast presenter running the slide show still tells the audience its slide.
 */
import React, { act } from 'react';
import { createRoot } from 'react-dom/client';
import type { Root } from 'react-dom/client';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import type { CollaborationContextValue } from './types';
import { useActiveSlidePresence } from './useActiveSlidePresence';

let root: Root;
let container: HTMLDivElement;

function Probe({
	collab,
	slide,
}: {
	collab: CollaborationContextValue | null;
	slide: number;
}): null {
	useActiveSlidePresence(collab, slide);
	return null;
}

beforeEach(() => {
	globalThis.IS_REACT_ACT_ENVIRONMENT = true;
	container = document.createElement('div');
	document.body.appendChild(container);
	root = createRoot(container);
});

afterEach(() => {
	act(() => root.unmount());
	container.remove();
	globalThis.IS_REACT_ACT_ENVIRONMENT = false;
});

describe('active slide presence', () => {
	it('publishes the active slide when the session starts and each time it changes', () => {
		const broadcastPresence = vi.fn();
		const collab = { broadcastPresence } as unknown as CollaborationContextValue;
		act(() => root.render(<Probe collab={collab} slide={0} />));
		expect(broadcastPresence).toHaveBeenLastCalledWith({ activeSlideIndex: 0 });

		act(() => root.render(<Probe collab={collab} slide={2} />));
		expect(broadcastPresence).toHaveBeenLastCalledWith({ activeSlideIndex: 2 });
		expect(broadcastPresence).toHaveBeenCalledTimes(2);
	});

	it('publishes nothing outside a session', () => {
		expect(() => act(() => root.render(<Probe collab={null} slide={1} />))).not.toThrow();
	});
});
