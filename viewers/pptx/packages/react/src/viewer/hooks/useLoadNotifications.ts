import type { CanvasSize } from 'ooxml-ui/pptx';
import { useEffect, useRef } from 'react';

/** What `onLoad` reports: the same detail the Svelte and vanilla bindings send. */
export interface ViewerLoadDetail {
	slideCount: number;
	canvasSize: CanvasSize;
}

export interface UseLoadNotificationsInput {
	loading: boolean;
	error: string | null;
	slideCount: number;
	canvasSize: CanvasSize;
	onLoad: ((detail: ViewerLoadDetail) => void) | undefined;
	onError: ((message: string) => void) | undefined;
}

/**
 * Host `onLoad` / `onError`: `onLoad` fires each time a load finishes without an error
 * (`loading` falling from true to false), `onError` once per distinct load failure message.
 * Callbacks and the loaded deck's numbers are read through refs, so a host that passes a new
 * closure every render does not re-fire either of them.
 */
export function useLoadNotifications(input: UseLoadNotificationsInput): void {
	const { loading, error } = input;
	const latest = useRef(input);
	latest.current = input;
	const wasLoading = useRef(loading);
	const announcedError = useRef<string | null>(null);

	useEffect(() => {
		const finished = wasLoading.current && !loading;
		wasLoading.current = loading;
		if (finished && !latest.current.error) {
			latest.current.onLoad?.({
				slideCount: latest.current.slideCount,
				canvasSize: latest.current.canvasSize,
			});
		}
	}, [loading]);

	useEffect(() => {
		if (error && error !== announcedError.current) {
			announcedError.current = error;
			latest.current.onError?.(error);
		} else if (!error) {
			announcedError.current = null;
		}
	}, [error]);
}
