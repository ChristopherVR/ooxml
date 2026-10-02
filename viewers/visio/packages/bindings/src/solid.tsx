import { createEffect, createSignal, onCleanup, onMount, type JSX } from 'solid-js';
import {
	mountFrameworkViewer,
	viewerHandle,
	viewerOptions,
	type MountedViewer,
	type ViewerHandle,
	type ViewerProps,
} from './common.js';
export interface VisioViewerProps extends ViewerProps {
	class?: string | undefined;
	style?: JSX.CSSProperties | undefined;
	viewerRef?: ((handle: ViewerHandle | undefined) => void) | undefined;
}
/** Native Solid host and ownership; all document work remains in the shared binding. */
export function VisioViewer(props: VisioViewerProps) {
	let host!: HTMLDivElement;
	const [binding, setBinding] = createSignal<MountedViewer>();
	let releasedRef: VisioViewerProps['viewerRef'];
	onMount(() => {
		setBinding(mountFrameworkViewer(host, viewerOptions(props)));
	});
	createEffect(() => {
		binding()?.update(viewerOptions(props));
	});
	createEffect(() => {
		const mounted = binding();
		const callback = props.viewerRef;
		if (!mounted || callback === releasedRef) return;
		const previous = releasedRef;
		releasedRef = callback;
		try {
			previous?.(undefined);
		} finally {
			callback?.(viewerHandle(binding));
		}
	});
	onCleanup(() => {
		const mounted = binding();
		setBinding(undefined);
		const release = releasedRef;
		releasedRef = undefined;
		try {
			release?.(undefined);
		} finally {
			mounted?.destroy();
		}
	});
	return <div ref={host} class={props.class} style={props.style} />;
}
export type {
	ViewerHandle,
	ViewerCallbacks,
	ViewerOptions,
	ViewerEvents,
	ViewerEditState,
	VsdxExportResult,
} from './common.js';
