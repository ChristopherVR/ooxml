import { createPresenterSessionController } from 'ooxml-ui/pptx';
import type {
	PresentationSnapshot,
	PresenterSessionController,
	PresenterSessionState,
} from 'ooxml-ui/pptx';

import { controllerState } from '../state/controller-state';
import type { ControllerState } from '../state/controller-state';

export interface PresenterSessionOptions {
	getSource: () => Uint8Array | ArrayBuffer | null | undefined;
	onAudienceSlide: (index: number) => void;
	onAudienceExit: () => void;
}

/**
 * The presenter and audience windows for the Svelte binding: the shared
 * `createPresenterSessionController` read through the store adapter.
 */
export class PresenterSession {
	readonly controller: PresenterSessionController;
	readonly #state: ControllerState<PresenterSessionState>;
	readonly isAudience: boolean;

	constructor(options: PresenterSessionOptions) {
		this.controller = createPresenterSessionController({
			getSource: options.getSource,
			onAudienceSlide: options.onAudienceSlide,
			onAudienceExit: options.onAudienceExit,
		});
		this.#state = controllerState(this.controller);
		this.isAudience = this.controller.getSnapshot().role === 'audience';
	}

	get audienceOpen(): boolean {
		return this.#state.current.audienceOpen;
	}
	get snapshot(): PresentationSnapshot {
		return this.#state.current.snapshot;
	}
	connect(): void {
		this.controller.connect();
	}
	openAudience(): boolean {
		return this.controller.openAudience();
	}
	swapDisplays(): Promise<boolean> {
		return this.controller.swapDisplays();
	}
	/** The presenter moved to `slideIndex`. */
	sync(slideIndex: number): void {
		if (!this.isAudience) {
			this.controller.updateSnapshot({ slideIndex });
		}
	}
	updateSnapshot(patch: Partial<PresentationSnapshot>): void {
		this.controller.updateSnapshot(patch);
	}
	closeAudience(): void {
		this.controller.closeAudience();
	}
	dispose(): void {
		this.controller.dispose();
	}
}
