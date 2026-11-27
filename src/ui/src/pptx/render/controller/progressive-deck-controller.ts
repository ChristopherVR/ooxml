/**
 * Editing a deck while the rest of it parses: the rules every binding follows
 * after a progressive open, as one headless controller.
 *
 * The parser delivers slides in presentation order (`onSlideLoaded`). A
 * binding that mounts the editor before the last one arrives stands each
 * unparsed slide in as a placeholder (`pendingSlide`) and reports every
 * arrival here. The rules that keep that safe:
 *
 * 1. An arrival only ever replaces its own placeholder, found by its stable
 *    load key, never a parsed slide, so no edit is overwritten. Slides the
 *    user moved keep their new position; a placeholder the user deleted stays
 *    deleted (and comes back filled if the deletion is undone).
 * 2. Edits to one slide run as soon as that slide is parsed. An edit aimed at
 *    a placeholder is blocked, never applied to a stale or empty model.
 * 3. Moving, deleting or inserting slides runs at once: placeholders move with
 *    the deck. Copying, duplicating or exporting slides needs their content,
 *    so it waits for those slides.
 * 4. Whole-deck work (save, export, print, collaboration, replace all, theme,
 *    slide size, apply to all) is queued until every slide is parsed and then
 *    runs in the order it was asked for. Save shows the parse progress meanwhile.
 * 5. Undo and redo restore snapshots taken while slides were pending; every
 *    restored slide list goes through `fillPendingSlides`, so a snapshot never
 *    brings a placeholder back over an arrived slide.
 * 6. Search covers the parsed slides and reports its coverage; a binding reruns
 *    it when the coverage grows. A slide show starts from parsed slides and waits
 *    at the first placeholder it reaches.
 * 7. A failed or superseded load drops the queue (nothing runs on a partial
 *    deck) and leaves the parsed slides as they are.
 *
 * @module render/controller/progressive-deck-controller
 */
import type { PptxSlide } from 'ooxml-core/pptx';

import { createControllerStore, type PptxController } from './controller';

const PENDING_PREFIX = 'pptx-pending-slide:';

/** A stand-in for the unparsed slide at load position `index`. */
export function pendingSlide(index: number): PptxSlide {
	return { id: `${PENDING_PREFIX}${index}`, rId: '', slideNumber: index + 1, elements: [] };
}

/** The load position a placeholder stands in for, or `undefined` for a parsed slide. */
export function pendingSlideIndex(slide: Pick<PptxSlide, 'id'> | undefined): number | undefined {
	if (!slide?.id.startsWith(PENDING_PREFIX)) return undefined;
	const index = Number(slide.id.slice(PENDING_PREFIX.length));
	return Number.isSafeInteger(index) && index >= 0 ? index : undefined;
}

/** Whether the slide is still a placeholder. */
export function isPendingSlide(slide: Pick<PptxSlide, 'id'> | undefined): boolean {
	return pendingSlideIndex(slide) !== undefined;
}

/** Replace every placeholder whose slide has arrived; the same array when none has. */
export function fillPendingSlides(
	slides: readonly PptxSlide[],
	arrived: ReadonlyMap<number, PptxSlide>,
): readonly PptxSlide[] {
	let changed = false;
	const next = slides.map((slide) => {
		const index = pendingSlideIndex(slide);
		const filled = index === undefined ? undefined : arrived.get(index);
		if (!filled) return slide;
		changed = true;
		return filled;
	});
	return changed ? next : slides;
}

/** An operation a binding asks about while the deck parses. */
export type ProgressiveOperation =
	| { kind: 'editSlide'; slide: Pick<PptxSlide, 'id'> | undefined }
	| { kind: 'arrangeSlides' }
	| { kind: 'copySlides'; slides: readonly (Pick<PptxSlide, 'id'> | undefined)[] }
	| { kind: 'deck' };

/** `run` now, `queue` until the deck has parsed, or `block` (the target is not parsed). */
export type ProgressiveDecision = 'run' | 'queue' | 'block';

/** Decide an operation by the rules in the module note. */
export function decideProgressiveOperation(
	operation: ProgressiveOperation,
	complete: boolean,
): ProgressiveDecision {
	if (complete) return 'run';
	switch (operation.kind) {
		case 'editSlide':
			return operation.slide && !isPendingSlide(operation.slide) ? 'run' : 'block';
		case 'arrangeSlides':
			return 'run';
		case 'copySlides':
			return operation.slides.every((slide) => slide && !isPendingSlide(slide)) ? 'run' : 'queue';
		case 'deck':
			return 'queue';
	}
}

/** The slides a search can read now, with their positions, and how much of the deck that is. */
export function progressiveSearchScope(slides: readonly PptxSlide[]) {
	const parsed = slides.flatMap((slide, index) =>
		isPendingSlide(slide) ? [] : [{ slide, index }],
	);
	return { slides: parsed, searched: parsed.length, total: slides.length };
}

/** Whether a slide show may step to `index` now, or waits for that slide to parse. */
export function progressiveShowStep(
	slides: readonly PptxSlide[],
	index: number,
): 'ready' | 'wait' | 'end' {
	const slide = slides[index];
	if (!slide) return 'end';
	return isPendingSlide(slide) ? 'wait' : 'ready';
}

export interface ProgressiveDeckState {
	/** Slides the parser has delivered. */
	parsed: number;
	/** Slides in the deck. */
	total: number;
	/** Every slide has parsed and the load finished. */
	complete: boolean;
	/** The load failed or was superseded: queued work was dropped. */
	failed: boolean;
	/** Whole-deck operations waiting for the parse, for the status text. */
	queued: number;
}

export interface ProgressiveDeckController extends PptxController<ProgressiveDeckState> {
	/** The slides that have arrived, by load position (for `fillPendingSlides`). */
	readonly arrived: ReadonlyMap<number, PptxSlide>;
	/** Record a parsed slide (from `onSlideLoaded`). */
	arrive(index: number, slide: PptxSlide): void;
	/** The load finished: run the queue in order. */
	finish(): void;
	/** The load failed or was superseded: drop the queue. */
	fail(): void;
	/** Run `task` now when the deck is complete, else when it completes. */
	whenComplete<T>(task: () => T | Promise<T>): Promise<T>;
}

/** One controller per load; create it with the slide count `onSlideLoaded` reports. */
export function createProgressiveDeckController(total: number): ProgressiveDeckController {
	const store = createControllerStore<ProgressiveDeckState>({
		parsed: 0,
		total,
		complete: total === 0,
		failed: false,
		queued: 0,
	});
	const arrived = new Map<number, PptxSlide>();
	let queue: Array<{ run: () => void; drop: (error: Error) => void }> = [];
	const drain = () => {
		const tasks = queue;
		queue = [];
		return tasks;
	};
	return {
		getSnapshot: store.getSnapshot,
		subscribe: store.subscribe,
		arrived,
		arrive(index, slide) {
			const state = store.getSnapshot();
			if (state.complete || state.failed || index < 0 || index >= state.total) return;
			if (arrived.has(index)) return;
			arrived.set(index, slide);
			store.set({ ...state, parsed: arrived.size });
		},
		finish() {
			const state = store.getSnapshot();
			if (state.complete || state.failed) return;
			const tasks = drain();
			store.set({ ...state, complete: true, queued: 0 });
			for (const task of tasks) task.run();
		},
		fail() {
			const state = store.getSnapshot();
			if (state.complete || state.failed) return;
			const tasks = drain();
			store.set({ ...state, failed: true, queued: 0 });
			const error = new Error('The presentation did not finish loading.');
			for (const task of tasks) task.drop(error);
		},
		whenComplete(task) {
			const state = store.getSnapshot();
			if (state.complete) return Promise.resolve().then(task);
			if (state.failed)
				return Promise.reject(new Error('The presentation did not finish loading.'));
			return new Promise((resolve, reject) => {
				queue.push({
					run: () => {
						Promise.resolve().then(task).then(resolve, reject);
					},
					drop: reject,
				});
				store.set({ ...store.getSnapshot(), queued: queue.length });
			});
		},
	};
}
