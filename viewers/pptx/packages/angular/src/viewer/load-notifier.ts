import type { CanvasSize } from 'ooxml-ui/pptx';

/** What the `load` output reports; identical in every binding. */
export interface ViewerLoadDetail {
	slideCount: number;
	canvasSize: CanvasSize;
}

/** The loader state the notifier reads on each change. */
export interface LoadSnapshot {
	loading: boolean;
	error: string | null;
	slideCount: number;
	canvasSize: CanvasSize;
}

/** Where the notifications go: the component's `load` and `error` outputs. */
export interface LoadSink {
	load(detail: ViewerLoadDetail): void;
	error(message: string): void;
}

/**
 * Turns loader state into the host's `load` / `error` events: `load` each time a load finishes
 * without an error (`loading` falling from true to false), `error` once per distinct failure.
 * Pure so the semantics are tested without a component; the component feeds it from an effect.
 */
export function createLoadNotifier(sink: LoadSink): (state: LoadSnapshot) => void {
	let wasLoading = false;
	let announcedError: string | null = null;
	return (state) => {
		if (wasLoading && !state.loading && !state.error) {
			sink.load({ slideCount: state.slideCount, canvasSize: state.canvasSize });
		}
		wasLoading = state.loading;
		if (state.error && state.error !== announcedError) {
			announcedError = state.error;
			sink.error(state.error);
		} else if (!state.error) {
			announcedError = null;
		}
	};
}
