import {
	DestroyRef,
	Injector,
	runInInjectionContext,
	ɵChangeDetectionScheduler as ChangeDetectionScheduler,
	ɵEffectScheduler as EffectScheduler,
} from '@angular/core';
import { TranslateService } from '@ngx-translate/core';
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { AutosaveService } from './autosave.service';
import { EditorStateService } from './editor-state.service';

vi.mock(import('ooxml-ui/pptx'), async () => {
	const actual = await vi.importActual<typeof import('ooxml-ui/pptx')>('ooxml-ui/pptx');
	return { ...actual, saveAutosaveSnapshot: vi.fn(async () => true) };
});

const { saveAutosaveSnapshot } = await import('ooxml-ui/pptx');

/**
 * The autosave TIMER contract, the Angular half of the polling pair.
 *
 * Angular and React poll on an interval while the document is dirty; Vue,
 * Svelte and Vanilla debounce on the slides signal being reassigned. A recovery
 * snapshot deliberately never clears the editor's dirty flag, so the polling
 * pair used to re-serialize and rewrite an identical deck on every tick, for as
 * long as the tab stayed open. These tests pin Angular onto the debounce
 * engines' trigger, and pin the cases where it must still write.
 *
 * Mirrors `packages/react/src/viewer/hooks/useAutosave.tick.test.tsx`; the two
 * must not drift.
 */
interface SchedulableEffect {
	run(): void;
}

/**
 * A bare injector rather than `TestBed`: this package has no Angular test
 * platform (no `@analogjs/vite-plugin-angular`), so `TestBed` cannot compile
 * its dynamic module. `bind()` registers an `effect`, which needs both
 * schedulers; the effect scheduler queues and is drained explicitly, because
 * running a watch inside `schedule()` is an error in Angular's own contract.
 */
function harness(): { autosave: AutosaveService; flushEffects: () => void } {
	const queued = new Set<SchedulableEffect>();
	const effects = {
		add: (effect: SchedulableEffect) => queued.add(effect),
		schedule: (effect: SchedulableEffect) => queued.add(effect),
		remove: (effect: SchedulableEffect) => queued.delete(effect),
		flush: () => {
			for (const effect of [...queued]) {
				queued.delete(effect);
				effect.run();
			}
		},
	};
	const injector = Injector.create({
		providers: [
			{ provide: TranslateService, useValue: { instant: (key: string) => key } },
			{ provide: DestroyRef, useValue: { onDestroy: () => () => {} } },
			{ provide: ChangeDetectionScheduler, useValue: { notify: () => {} } },
			{ provide: EffectScheduler, useValue: effects },
		],
	});
	return {
		autosave: runInInjectionContext(injector, () => new AutosaveService()),
		flushEffects: () => effects.flush(),
	};
}

describe('autosave timer redundancy', () => {
	beforeEach(() => {
		vi.useFakeTimers();
		vi.mocked(saveAutosaveSnapshot).mockClear();
	});

	afterEach(() => {
		vi.useRealTimers();
	});

	interface Bound {
		serialize: ReturnType<typeof vi.fn>;
		sources: { value: readonly unknown[] };
		autosave: AutosaveService;
		tick: () => Promise<void>;
	}

	function bind(withSources: boolean): Bound {
		const serialize = vi.fn(async () => new Uint8Array([1, 2, 3]));
		const sources: { value: readonly unknown[] } = { value: [] };
		const { autosave, flushEffects } = harness();
		autosave.bind({
			enabled: () => true,
			filePath: () => 'deck.pptx',
			isDirty: () => true,
			serialize,
			intervalMs: () => 10_000,
			...(withSources ? { changeSources: () => sources.value } : {}),
		});
		// The interval is armed inside the bind effect.
		flushEffects();
		return {
			serialize,
			sources,
			autosave,
			tick: async () => {
				await vi.advanceTimersByTimeAsync(10_000);
			},
		};
	}

	it('writes the first snapshot, then skips ticks that change nothing', async () => {
		const slides = [{ id: 'slide1' }];
		const bound = bind(true);
		bound.sources.value = [slides];

		await bound.tick();
		expect(bound.serialize).toHaveBeenCalledOnce();
		expect(saveAutosaveSnapshot).toHaveBeenCalledOnce();

		await bound.tick();
		await bound.tick();
		expect(bound.serialize).toHaveBeenCalledOnce();
		expect(saveAutosaveSnapshot).toHaveBeenCalledOnce();
	});

	it('writes again as soon as an edit reassigns the slides', async () => {
		const slides = [{ id: 'slide1' }];
		const bound = bind(true);
		bound.sources.value = [slides];
		await bound.tick();
		expect(bound.serialize).toHaveBeenCalledOnce();

		// An immutable edit: same content, new array.
		bound.sources.value = [[...slides]];
		await bound.tick();
		expect(bound.serialize).toHaveBeenCalledTimes(2);
	});

	it('writes on every tick when the host supplies no change sources', async () => {
		const bound = bind(false);
		await bound.tick();
		await bound.tick();
		expect(bound.serialize).toHaveBeenCalledTimes(2);
	});

	it('honours an explicit host cadence instead of the AutoRecover default', async () => {
		const serialize = vi.fn(async () => new Uint8Array([1, 2, 3]));
		const { autosave, flushEffects } = harness();
		autosave.bind({
			enabled: () => true,
			filePath: () => 'deck.pptx',
			isDirty: () => true,
			serialize,
			// A host asking for a two second cadence is stating a policy; the old
			// seconds path clamped anything under ten seconds back up to ten.
			intervalMs: () => 2000,
		});
		flushEffects();
		await vi.advanceTimersByTimeAsync(2000);
		expect(serialize).toHaveBeenCalledOnce();
	});

	it('never suppresses an explicit triggerAutosave', async () => {
		const bound = bind(true);
		bound.sources.value = [[{ id: 'slide1' }]];

		await bound.autosave.triggerAutosave();
		await bound.autosave.triggerAutosave();
		expect(bound.serialize).toHaveBeenCalledTimes(2);
	});

	it('does not mistake an edit during serialization for an already saved source', async () => {
		const bound = bind(true);
		let finish!: (data: Uint8Array) => void;
		bound.serialize.mockImplementationOnce(
			() =>
				new Promise<Uint8Array>((resolve) => {
					finish = resolve;
				}),
		);
		bound.sources.value = [[{ id: 'before' }]];
		const pending = bound.autosave.triggerAutosave();
		bound.sources.value = [[{ id: 'after' }]];
		finish(new Uint8Array([1]));
		await pending;
		expect(saveAutosaveSnapshot).toHaveBeenCalledOnce();
		await bound.tick();
		expect(bound.serialize).toHaveBeenCalledTimes(2);
		expect(saveAutosaveSnapshot).toHaveBeenCalledTimes(2);
		await bound.tick();
		expect(bound.serialize).toHaveBeenCalledTimes(2);
	});
});

/**
 * The edit the browser recovery specs drive (Home > New Slide) must reach the
 * snapshot writer through the REAL editor state, not a stubbed dirty flag.
 *
 * Every test above binds `isDirty: () => true`, so none of them could notice an
 * editor that never raises the flag, which is exactly how React once shipped
 * with no crash recovery at all. `autosave-recovery-prompt.spec.ts` and
 * `autosave-recovery-encryption.spec.ts` used to skip any binding that wrote no
 * snapshot after New Slide; this pins the Angular half of that guarantee.
 */
describe('a ribbon New Slide on a freshly loaded deck', () => {
	beforeEach(() => {
		vi.useFakeTimers();
		vi.mocked(saveAutosaveSnapshot).mockClear();
	});

	afterEach(() => {
		vi.useRealTimers();
	});

	it.each([
		{ layoutPath: undefined, label: 'the plain button' },
		{ layoutPath: 'ppt/slideLayouts/slideLayout2.xml', label: 'a layout from its menu' },
	])('writes a recovery snapshot after $label', async ({ layoutPath }) => {
		const editor = new EditorStateService();
		editor.setSlides([{ id: 's1', rId: 's1', slideNumber: 1, elements: [] }]);
		const serialize = vi.fn(async () => new Uint8Array([1, 2, 3]));
		const { autosave, flushEffects } = harness();
		autosave.bind({
			enabled: () => true,
			filePath: () => 'deck.pptx',
			isDirty: () => editor.dirty(),
			serialize,
			intervalMs: () => 2000,
			changeSources: () => [editor.slides(), editor.templateElementsBySlideId()],
		});
		flushEffects();

		// Opening a deck is not an edit: nothing to recover yet.
		await vi.advanceTimersByTimeAsync(2000);
		expect(saveAutosaveSnapshot).not.toHaveBeenCalled();

		// The handler `ribbon-home-section` runs for `home.slides.newSlide`.
		editor.addSlide(0, layoutPath);
		expect(editor.dirty()).toBeTruthy();

		await vi.advanceTimersByTimeAsync(2000);
		expect(serialize).toHaveBeenCalledOnce();
		expect(saveAutosaveSnapshot).toHaveBeenCalledWith('deck.pptx', new Uint8Array([1, 2, 3]));
	});
});
