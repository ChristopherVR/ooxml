import {
	createElement,
	forwardRef,
	useEffect,
	useImperativeHandle,
	useRef,
	type CSSProperties,
} from 'react';
import {
	mountFrameworkViewer,
	viewerHandle,
	viewerOptions,
	type MountedViewer,
	type ViewerHandle,
	type ViewerProps,
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
export type {
	ViewerHandle,
	ViewerCallbacks,
	ViewerOptions,
	ViewerEvents,
	ViewerEditState,
	VsdxExportResult,
} from './common.js';
