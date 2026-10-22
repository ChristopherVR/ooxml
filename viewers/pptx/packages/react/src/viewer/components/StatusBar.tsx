import { dispatchStatusBarIntent, statusBarState, statusBarViewMode } from 'ooxml-ui/pptx';
import type { PptxUiStatusBarElement } from 'ooxml-ui/pptx';
import React, { useEffect, useRef } from 'react';
import { useTranslation } from 'react-i18next';

import type { AutosaveStatus } from '../hooks/useAutosave';

export interface StatusBarProps {
	slideCount: number;
	activeSlideIndex: number;
	isDirty: boolean;
	autosaveStatus?: AutosaveStatus;
	/** Current zoom scale (0-1+ range, e.g. 1.0 = 100%). */
	scale?: number;
	/** Callback to zoom in. */
	onZoomIn?: () => void;
	/** Callback to zoom out. */
	onZoomOut?: () => void;
	/** Callback to zoom to fit. */
	onZoomToFit?: () => void;
	/** Whether the notes panel is expanded. */
	isNotesExpanded?: boolean;
	/** Toggle the notes panel. */
	onToggleNotes?: () => void;
	/** Current viewer mode. */
	mode?: string;
	/** Callback to switch viewer mode. */
	onSetMode?: (mode: 'edit' | 'present') => void;
	/** Callback to toggle slide sorter view. */
	onToggleSlideSorter?: () => void;
	/** Optional collaboration status indicator rendered inline. */
	collaborationSlot?: React.ReactNode;
	/** Hides the zoom in/out/to-fit control cluster. Maps to `hiddenActions: ['zoom']`. */
	hideZoomControls?: boolean;
	/** Hides the notes-panel toggle button. Maps to `hiddenActions: ['notes']`. */
	hideNotesToggle?: boolean;
	/** Hides the quick "Slide Show" (fullscreen present) toggle button. Maps to `hiddenActions: ['fullscreen']`. */
	hideFullscreenToggle?: boolean;
}

/**
 * Thin adapter around the shared `pptx-ui-status-bar`: it maps viewer state to the
 * element's controlled state and routes `status-request` intents to the handlers.
 */
export function StatusBar(p: StatusBarProps): React.ReactElement {
	const { t } = useTranslation();
	const ref = useRef<PptxUiStatusBarElement>(null);
	useEffect(() => {
		const host = ref.current;
		if (!host) {
			return;
		}
		host.state = statusBarState({
			slideCount: p.slideCount,
			activeSlideIndex: p.activeSlideIndex,
			dirty: p.isDirty,
			autosave: p.autosaveStatus,
			zoomPercent: p.scale !== undefined && !p.hideZoomControls ? p.scale * 100 : undefined,
			notesExpanded: p.isNotesExpanded,
			viewMode: statusBarViewMode(p.mode),
			handles: {
				notes: Boolean(p.onToggleNotes) && !p.hideNotesToggle,
				viewModes: Boolean(p.onSetMode),
				sorter: Boolean(p.onToggleSlideSorter),
				slideShow: !p.hideFullscreenToggle,
			},
			translate: t,
		});
	}, [p, t]);
	useEffect(() => {
		const host = ref.current;
		if (!host) {
			return;
		}
		const request = (event: Event) =>
			dispatchStatusBarIntent(event, {
				notes: p.onToggleNotes,
				normal: () => p.onSetMode?.('edit'),
				sorter: p.onToggleSlideSorter,
				slideShow: () => p.onSetMode?.('present'),
				zoomOut: p.onZoomOut,
				zoomFit: p.onZoomToFit,
				zoomIn: p.onZoomIn,
			});
		host.addEventListener('status-request', request);
		return () => host.removeEventListener('status-request', request);
	}, [p]);
	return (
		<pptx-ui-status-bar ref={ref}>
			{p.collaborationSlot ? (
				<div slot='collaboration' style={{ display: 'contents' }}>
					{p.collaborationSlot}
				</div>
			) : null}
		</pptx-ui-status-bar>
	);
}
