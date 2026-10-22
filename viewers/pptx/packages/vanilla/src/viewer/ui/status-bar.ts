import { dispatchStatusBarIntent, registerPptxWebControls, statusBarState } from 'ooxml-ui/pptx';
import type { StatusBarAutosave, ToolbarActionId } from 'ooxml-ui/pptx';

import type { Translator } from '../i18n';
import type { RibbonNavState } from './ribbon/ribbon-types';

export interface StatusBarHandlers {
	toggleNotes(): void;
	/** Return to the normal editing view (exit presentation + close slide sorter). */
	normalView(): void;
	openSlideSorter(): void;
	togglePresentation(): void;
	zoomIn(): void;
	zoomOut(): void;
	zoomToFit(): void;
}

export interface StatusBar {
	el: HTMLElement;
	/** Reflect the slide counter and zoom percent. */
	update(state: RibbonNavState): void;
	setNotesExpanded(expanded: boolean): void;
	/** Reflect the autosave engine (undefined when it is off); the dirty flag is the fallback. */
	setAutosave(autosave: StatusBarAutosave | undefined): void;
	/** Reflect the unsaved-changes flag in the idle save-state text. */
	setDirty(dirty: boolean): void;
	/** Reflect presentation state so the Normal / Slide Show buttons show the
	    active view (derived, not hardcoded). */
	setPresenting(presenting: boolean): void;
}

/**
 * PowerPoint-style status bar docked below the notes strip. A thin adapter over
 * the shared `pptx-ui-status-bar`: slide counter, language and save state on the
 * left; notes toggle, view buttons and zoom on the right.
 */
export function createStatusBar(
	doc: Document,
	t: Translator,
	handlers: StatusBarHandlers,
	hiddenActions?: readonly ToolbarActionId[],
	/**
	 * Optional collaboration-status element, projected between the view-mode
	 * cluster and the zoom cluster (React parity's `collaborationSlot`). Hosts
	 * can supply their own connection-status indicator here.
	 */
	collaborationSlot?: HTMLElement,
): StatusBar {
	registerPptxWebControls();
	const el = doc.createElement('pptx-ui-status-bar');
	el.className = 'pptxv-statusbar';
	if (collaborationSlot) {
		collaborationSlot.slot = 'collaboration';
		el.append(collaborationSlot);
	}
	let dirty = false;
	let autosave: StatusBarAutosave | undefined;
	let nav = { current: 0, total: 0, zoomPercent: 100 };
	let notesExpanded = false;
	let presenting = false;

	const sync = (): void => {
		el.state = statusBarState({
			slideCount: nav.total,
			activeSlideIndex: nav.current,
			dirty,
			autosave,
			zoomPercent: nav.zoomPercent,
			notesExpanded,
			// The slide-sorter overlay is modal and self-dismissing, so it is not
			// tracked as a persistent active view.
			viewMode: presenting ? 'slideShow' : 'normal',
			hiddenActions,
			translate: t,
		});
	};
	el.addEventListener('status-request', (event) =>
		dispatchStatusBarIntent(event, {
			notes: handlers.toggleNotes,
			normal: handlers.normalView,
			sorter: handlers.openSlideSorter,
			slideShow: handlers.togglePresentation,
			zoomOut: handlers.zoomOut,
			zoomFit: handlers.zoomToFit,
			zoomIn: handlers.zoomIn,
		}),
	);
	sync();

	return {
		el,
		update(state) {
			nav = { current: state.current, total: state.total, zoomPercent: state.zoomPercent };
			sync();
		},
		setNotesExpanded(expanded) {
			notesExpanded = expanded;
			sync();
		},
		setAutosave(next) {
			autosave = next;
			sync();
		},
		setDirty(isDirty) {
			dirty = isDirty;
			sync();
		},
		setPresenting(next) {
			presenting = next;
			sync();
		},
	};
}
