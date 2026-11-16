import { applyRehearsalTimings, createPresentationRehearsalController } from 'ooxml-ui/pptx';
import type { PresentationRehearsalController } from 'ooxml-ui/pptx';

import type { EditorState } from '../editor/editor-state.svelte';
import { controllerState } from '../state/controller-state';

/**
 * Rehearse Timings for the Svelte binding: the shared
 * `createPresentationRehearsalController` read through the store adapter. The
 * controller owns the clocks and their tick; this class only names its state
 * for the templates and writes the saved timings through the editor.
 */
export class RehearseState {
	readonly controller: PresentationRehearsalController = createPresentationRehearsalController();
	readonly #state = controllerState(this.controller);

	get active(): boolean {
		return this.#state.current.active;
	}
	get paused(): boolean {
		return this.#state.current.paused;
	}
	get summaryOpen(): boolean {
		return this.#state.current.summaryOpen;
	}
	get current(): number {
		return this.#state.current.slideIndex;
	}
	get elapsedMs(): number {
		return this.#state.current.slideElapsedMs;
	}
	get times(): Map<number, number> {
		return new Map(
			Object.entries(this.#state.current.timings).map(([index, ms]) => [Number(index), ms]),
		);
	}
	get totalMs(): number {
		return this.#state.current.recordedTotalMs;
	}
	start(index: number): void {
		this.controller.start(index);
	}
	move(index: number): void {
		this.controller.moveTo(index);
	}
	togglePause(): void {
		this.controller.togglePause();
	}
	finish(): void {
		this.controller.finish();
	}
	save(editor: EditorState): void {
		editor.commitSlides(applyRehearsalTimings(editor.slides, this.controller.save()));
	}
	discard(): void {
		this.controller.discard();
	}
}
