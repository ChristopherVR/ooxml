import {
	DestroyRef,
	Injector,
	signal,
	runInInjectionContext,
	ɵChangeDetectionScheduler as ChangeDetectionScheduler,
	ɵEffectScheduler as EffectScheduler,
} from '@angular/core';
import type { WritableSignal } from '@angular/core';
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

describe('the shared autosave engine through AutosaveService', () => {
	beforeEach(() => {
		vi.useFakeTimers();
		vi.mocked(saveAutosaveSnapshot).mockClear();
	});

	afterEach(() => {
		vi.useRealTimers();
	});

	interface Bound {
		serialize: ReturnType<typeof vi.fn>;
		sources: WritableSignal<readonly unknown[]>;
		autosave: AutosaveService;
		flushEffects: () => void;
		tick: () => Promise<void>;
	}

	function bind(withSources: boolean): Bound {
		const serialize = vi.fn(async () => new Uint8Array([1, 2, 3]));
		const sources = signal<readonly unknown[]>([[{ id: 'slide1' }]]);
		const { autosave, flushEffects } = harness();
		autosave.bind({
			enabled: () => true,
			filePath: () => 'deck.pptx',
			isDirty: () => true,
			serialize,
			intervalMs: () => 10_000,
			...(withSources ? { changeSources: () => sources() } : {}),
		});
		flushEffects();
		return {
			serialize,
			sources,
			autosave,
			flushEffects,
			tick: async () => {
				await vi.advanceTimersByTimeAsync(10_000);
			},
		};
	}

	it('hands the written snapshot bytes to the host (autosaved)', async () => {
		const saved = vi.fn();
		const bytes = new Uint8Array([4, 5, 6]);
		const { autosave, flushEffects } = harness();
		autosave.bind({
			enabled: () => true,
			filePath: () => 'deck.pptx',
			isDirty: () => true,
			serialize: async () => bytes,
			intervalMs: () => 10_000,
			saved,
		});
		flushEffects();

		await vi.advanceTimersByTimeAsync(10_000);

		expect(saveAutosaveSnapshot).toHaveBeenCalledWith('deck.pptx', bytes, undefined);
		expect(saved).toHaveBeenCalledExactlyOnceWith(bytes);
		expect(autosave.status()).toMatchObject({ state: 'saved' });
	});

	it('writes the first snapshot, then nothing while the deck is unchanged', async () => {
		const bound = bind(true);
		await bound.tick();
		expect(saveAutosaveSnapshot).toHaveBeenCalledOnce();
		await bound.tick();
		await bound.tick();
		expect(bound.serialize).toHaveBeenCalledOnce();
	});

	it('writes again one interval after an edit reassigns the slides', async () => {
		const bound = bind(true);
		await bound.tick();
		bound.sources.set([[{ id: 'slide1' }]]);
		bound.flushEffects();
		await bound.tick();
		expect(bound.serialize).toHaveBeenCalledTimes(2);
	});

	it('without change sources, the clean-to-dirty transition is the one edit', async () => {
		const bound = bind(false);
		await bound.tick();
		await bound.tick();
		expect(bound.serialize).toHaveBeenCalledOnce();
	});

	it('honours an explicit host cadence instead of the AutoRecover default', async () => {
		const serialize = vi.fn(async () => new Uint8Array([1, 2, 3]));
		const { autosave, flushEffects } = harness();
		autosave.bind({
			enabled: () => true,
			filePath: () => 'deck.pptx',
			isDirty: () => true,
			serialize,
			intervalMs: () => 2000,
		});
		flushEffects();
		await vi.advanceTimersByTimeAsync(2000);
		expect(serialize).toHaveBeenCalledOnce();
	});

	it('never suppresses an explicit triggerAutosave', async () => {
		const bound = bind(true);
		await bound.autosave.triggerAutosave();
		await bound.autosave.triggerAutosave();
		expect(bound.serialize).toHaveBeenCalledTimes(2);
	});

	it('keeps an edit made during a write for the next interval', async () => {
		const bound = bind(true);
		let finish!: (data: Uint8Array) => void;
		bound.serialize.mockImplementationOnce(
			() =>
				new Promise<Uint8Array>((resolve) => {
					finish = resolve;
				}),
		);
		await bound.tick();
		bound.sources.set([[{ id: 'after' }]]);
		bound.flushEffects();
		finish(new Uint8Array([1]));
		await vi.advanceTimersByTimeAsync(0);
		expect(saveAutosaveSnapshot).toHaveBeenCalledOnce();
		await bound.tick();
		expect(saveAutosaveSnapshot).toHaveBeenCalledTimes(2);
		await bound.tick();
		expect(bound.serialize).toHaveBeenCalledTimes(2);
	});

	it('reports the shared disabled status when autosave is off', () => {
		const { autosave, flushEffects } = harness();
		autosave.bind({
			enabled: () => false,
			disabledReason: () => 'autosave_host_off',
			filePath: () => 'deck.pptx',
			isDirty: () => false,
			serialize: async () => null,
		});
		flushEffects();
		expect(autosave.status()).toEqual({ state: 'disabled', reason: 'autosave_host_off' });
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
		flushEffects();

		await vi.advanceTimersByTimeAsync(2000);
		expect(serialize).toHaveBeenCalledOnce();
		expect(saveAutosaveSnapshot).toHaveBeenCalledWith(
			'deck.pptx',
			new Uint8Array([1, 2, 3]),
			undefined,
		);
	});
});

describe('in-progress inline text', () => {
	beforeEach(() => {
		vi.useFakeTimers();
		vi.mocked(saveAutosaveSnapshot).mockClear();
	});

	afterEach(() => {
		vi.useRealTimers();
	});

	it('counts uncommitted typing on a clean deck as an edit and hands it to the serialiser', async () => {
		const serialize = vi.fn(async () => new Uint8Array([1]));
		const { autosave, flushEffects } = harness();
		autosave.bind({
			enabled: () => true,
			filePath: () => 'deck.pptx',
			isDirty: () => false,
			serialize,
			intervalMs: () => 2000,
			inlineEditingId: () => 'title',
		});
		flushEffects();
		autosave.noteInlineDraft({ slideId: 's1', elementId: 'title', text: 'Typed' });
		await vi.advanceTimersByTimeAsync(2000);
		expect(serialize).toHaveBeenCalledExactlyOnceWith({
			slideId: 's1',
			elementId: 'title',
			text: 'Typed',
		});
	});
});
