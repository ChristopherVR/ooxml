/**
 * A Design gallery hover preview swaps previewed slides into the live `slides` ref. It must not
 * mark the document dirty, arm autosave, or end up in a crash-recovery snapshot.
 */
// @vitest-environment node
import { PptxHandler } from 'pptx-viewer-core';
import type { PptxSlide } from 'pptx-viewer-core';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { effectScope, nextTick, ref, shallowRef } from 'vue';

import { useAutosave } from './useAutosave';
import { useLoadContent } from './useLoadContent';

const slide = (id: string): PptxSlide => ({ id, elements: [] }) as unknown as PptxSlide;

describe('useAutosave while a gallery preview shows', () => {
	beforeEach(() => vi.useFakeTimers());
	afterEach(() => vi.useRealTimers());

	it('neither marks the document dirty nor saves, and treats the restore as no edit', async () => {
		const original = [slide('a')];
		const slides = ref<PptxSlide[]>(original);
		const paused = ref(false);
		const onSave = vi.fn(() => Promise.resolve());
		const scope = effectScope();
		const api = scope.run(() => useAutosave({ slides, intervalMs: 1000, onSave, paused }))!;

		paused.value = true;
		slides.value = [slide('preview')];
		slides.value = original;
		paused.value = false;
		await vi.advanceTimersByTimeAsync(5000);

		expect(api.isDirty.value).toBeFalsy();
		expect(onSave).not.toHaveBeenCalled();

		// A real edit afterwards still counts.
		slides.value = [slide('a'), slide('b')];
		expect(api.isDirty.value).toBeTruthy();
		scope.stop();
	});
});

describe('the recovery snapshot while a gallery preview shows', () => {
	it('serializes the real deck, not the previewed one', async () => {
		const { handler, data } = await PptxHandler.create({ initialSlideCount: 2 });
		const bytes = await handler.save(data.slides);
		handler.dispose();
		const scope = effectScope();
		try {
			await scope.run(async () => {
				const content = ref<Uint8Array | null>(bytes);
				const held = shallowRef<
					{ slides: PptxSlide[]; templateElementsBySlideId: Record<string, never[]> } | undefined
				>(undefined);
				const deck = useLoadContent(() => content.value, { getHeldDeck: () => held.value });
				for (let i = 0; i < 200 && (deck.loading.value || !deck.handler.value); i++) {
					await nextTick();
					await new Promise((resolve) => setTimeout(resolve, 10));
				}
				const real = deck.slides.value;
				expect(real).toHaveLength(2);
				// The preview shows: one slide fewer in the live ref, the real deck held aside.
				held.value = { slides: real, templateElementsBySlideId: {} };
				deck.slides.value = real.slice(0, 1);

				const snapshot = await deck.getRecoverySnapshot();
				const reopened = new PptxHandler();
				try {
					expect((await reopened.load(snapshot.slice().buffer as ArrayBuffer)).slides).toHaveLength(
						2,
					);
				} finally {
					reopened.dispose();
				}
			});
		} finally {
			scope.stop();
		}
	}, 120_000);
});
