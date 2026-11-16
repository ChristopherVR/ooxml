// @vitest-environment jsdom
import { afterEach, expect, it, vi } from 'vitest';

afterEach(() => {
	vi.useRealTimers();
});

const delivered = (raf: typeof requestAnimationFrame) =>
	Promise.race([
		new Promise((resolve) => raf(() => resolve(true))),
		new Promise((resolve) => setTimeout(() => resolve(false), 500)),
	]);

it('keeps animation frames running after a test fakes the timers around a pending frame', async () => {
	// The shared setup's frame (dom-hygiene.ts); code that kept a reference before the timers were
	// faked still calls it while they are.
	const raf = window.requestAnimationFrame;
	const realTimeout = setTimeout;
	raf(() => undefined);
	vi.useFakeTimers();
	// A real frame tick lands while time is fake, then a frame is asked for in fake time.
	await new Promise((resolve) => realTimeout(resolve, 50));
	raf(() => undefined);
	vi.useRealTimers();
	expect(await delivered(raf)).toBe(true);
	expect(await delivered(requestAnimationFrame)).toBe(true);
});
