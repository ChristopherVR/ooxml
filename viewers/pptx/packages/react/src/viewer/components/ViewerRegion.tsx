/**
 * ViewerRegion: hosts one host-supplied region renderer (`customization.regions`
 * and `backstage.extraPages`) in the place of a built-in panel, dialog body or
 * File tab page.
 *
 * The mount, update and teardown semantics all live in the shared
 * `ViewerRegionHost`; this component only owns a container `div`, creates the
 * host when it mounts, hands it the renderer and the freshly built context
 * whenever either changes, and destroys it on unmount.
 */
import type {
	HostBackstageRegionId,
	PowerPointViewerAPI,
	ViewerMode,
	ViewerRegionContext,
	ViewerRegionId,
	ViewerRegionRenderer,
} from 'ooxml-ui/pptx';
import { ViewerRegionHost, getRegionRenderer } from 'ooxml-ui/pptx';
import { createContext, useContext, useLayoutEffect, useRef } from 'react';
import type { CSSProperties, ReactElement } from 'react';

import { useViewerCustomizationContext } from './viewer-customization-context';

/** The part of a region context that comes from viewer state (not per region). */
export interface ViewerRegionBase {
	readonly viewer: PowerPointViewerAPI;
	readonly mode: ViewerMode;
	readonly canEdit: boolean;
	readonly activeSlideIndex: number;
	readonly slideCount: number;
	readonly selectedElementIds: readonly string[];
	readonly locale: string;
}

/** Provided by `PowerPointViewer`; `null` when a component is rendered standalone. */
export const ViewerRegionBaseContext = /* @__PURE__ */ createContext<ViewerRegionBase | null>(null);

/**
 * The host renderer replacing `id`, or `undefined` when none is set OR no region
 * context can be built (rendered outside `PowerPointViewer`): the site then shows
 * the built-in instead of an empty body.
 */
export function useRegionRenderer(id: ViewerRegionId): ViewerRegionRenderer | undefined {
	const customization = useViewerCustomizationContext();
	const base = useContext(ViewerRegionBaseContext);
	return base ? getRegionRenderer(customization, id) : undefined;
}

export interface ViewerRegionProps {
	region: ViewerRegionId | HostBackstageRegionId;
	renderer: ViewerRegionRenderer;
	/** Dismisses the region (closes the dialog, leaves the File tab, ...). Default: no-op. */
	close?: () => void;
	/** Extra context fields of one region (the Options dialog's `options`, `setOption`, ...). */
	extra?: Record<string, unknown>;
	className?: string;
	style?: CSSProperties;
}

const FILL: CSSProperties = { width: '100%', height: '100%' };
const NOOP = (): void => {};

export function ViewerRegion(props: ViewerRegionProps): ReactElement {
	const { region, renderer, close = NOOP, extra, className, style } = props;
	const base = useContext(ViewerRegionBaseContext);
	const containerRef = useRef<HTMLDivElement>(null);
	const hostRef = useRef<ViewerRegionHost | null>(null);

	useLayoutEffect(() => {
		const host = new ViewerRegionHost(containerRef.current!);
		hostRef.current = host;
		return () => {
			host.destroy();
			hostRef.current = null;
		};
	}, []);

	// Deliberately runs after every render: the host compares renderer and
	// context itself and only remounts or updates on a real change.
	useLayoutEffect(() => {
		if (!base) {
			return;
		}
		const context = { ...base, ...extra, region, close } as ViewerRegionContext;
		hostRef.current?.render(renderer, context);
	});

	return (
		<div
			ref={containerRef}
			className={className ? `pptx-region-host ${className}` : 'pptx-region-host'}
			style={style ?? FILL}
		/>
	);
}
