import {
	createElement,
	forwardRef,
	useEffect,
	useImperativeHandle,
	useMemo,
	useRef,
	useSyncExternalStore,
	type CSSProperties,
	type RefObject,
} from 'react';
import {
	mountFrameworkViewer,
	viewerHandle,
	viewerOptions,
	viewerStateSource,
	type MountedViewer,
	type ViewerHandle,
	type ViewerProps,
	type ViewerState,
} from './common.js';
export interface VisioViewerProps extends ViewerProps {
	className?: string | undefined;
	style?: CSSProperties | undefined;
	'aria-label'?: string | undefined;
}
/** React owns only the host and lifecycle; the shared element owns all viewer behavior. */
export const VisioViewer = forwardRef<ViewerHandle, VisioViewerProps>(
	function VisioViewer(props, ref) {
		const host = useRef<HTMLDivElement>(null);
		const binding = useRef<MountedViewer | undefined>(undefined);
		useEffect(() => {
			const mounted = mountFrameworkViewer(host.current!, viewerOptions(props));
			binding.current = mounted;
			return () => {
				binding.current = undefined;
				mounted.destroy();
			};
		}, []);
		useEffect(() => {
			binding.current?.update(viewerOptions(props));
		});
		useImperativeHandle(ref, () => viewerHandle(() => binding.current), []);
		return createElement('div', {
			ref: host,
			className: props.className,
			style: props.style,
			'aria-label': props['aria-label'],
		});
	},
);
/**
 * Reactive viewer state (page, zoom, selection, search, edit and undo state) for the `ref`
 * given to `<VisioViewer>`, through React's external-store hook. `null` until mounted.
 */
export function useVisioViewerState(
	ref: RefObject<ViewerHandle | null | undefined>,
): ViewerState | null {
	const source = useMemo(() => viewerStateSource(() => ref.current), [ref]);
	return useSyncExternalStore(source.subscribe, source.getSnapshot, () => null);
}
export type {
	ViewerState,
	ViewerHandle,
	ViewerCallbacks,
	ViewerOptions,
	ViewerEvents,
	ViewerEditState,
	VsdxExportResult,
} from './common.js';
