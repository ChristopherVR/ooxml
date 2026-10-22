/**
 * Shared model for the bottom status bar (`pptx-ui-status-bar`).
 *
 * The slide counter, save indicator, notes toggle, view-mode buttons and zoom
 * cluster were hand-built five times. The view owns the markup, pressed state
 * and gating; hosts own every effect (notes panel, view switching, zoom).
 * {@link statusBarState} turns a host's own facts into the view's state and
 * {@link dispatchStatusBarIntent} routes the view's requests, so the gating and
 * the save rule are written once.
 */
import { isActionHidden } from './toolbar-actions';
import type { ToolbarActionId } from './toolbar-actions';

export type StatusBarTranslate = (key: string, params?: Record<string, string | number>) => string;

/** Public intent ids. Hosts map them to their native handlers. */
export type StatusBarControlId =
	| 'notes'
	| 'normal'
	| 'sorter'
	| 'slideShow'
	| 'zoomOut'
	| 'zoomFit'
	| 'zoomIn';

export type StatusBarSaveKind = 'idle' | 'saving' | 'error';
export type StatusBarViewMode = 'normal' | 'sorter' | 'slideShow';

export interface StatusBarViewState {
	slideCount: number;
	/** Zero-based active slide; clamped into range for display. */
	activeSlideIndex: number;
	/** Already-translated save indicator, see {@link resolveStatusBarSave}. */
	saveText: string;
	saveKind?: StatusBarSaveKind;
	/** Omit to hide the whole zoom cluster. */
	zoomPercent?: number;
	showNotes?: boolean;
	notesExpanded?: boolean;
	/** Default true. False drops the Normal / Sorter / Slide Show cluster. */
	showViewModes?: boolean;
	/** Default true. */
	showSorter?: boolean;
	/** Default true. */
	showSlideShow?: boolean;
	/** The pressed view button; omit when none is active. */
	viewMode?: StatusBarViewMode;
	translate?: StatusBarTranslate;
}

export type StatusBarIntent = { id: StatusBarControlId };

export interface StatusBarAutosave {
	state: 'idle' | 'disabled' | 'saving' | 'saved' | 'error';
	/** Epoch ms of the last successful save; "saved" without it reads as dirty/all-saved. */
	timestamp?: number;
}

function formatAge(timestamp: number, t: StatusBarTranslate, now: number): string {
	const minutes = Math.floor((now - timestamp) / 60_000);
	if (minutes < 1) {
		return t('pptx.autosave.justNow');
	}
	if (minutes === 1) {
		return t('pptx.autosave.oneMinAgo');
	}
	return t('pptx.autosave.minutesAgo', { count: minutes });
}

/** The one save-indicator rule: autosave state first, then the dirty flag. */
export function resolveStatusBarSave(
	t: StatusBarTranslate,
	autosave: StatusBarAutosave | undefined,
	dirty: boolean,
	now: number = Date.now(),
): { text: string; kind: StatusBarSaveKind } {
	if (autosave?.state === 'saving') {
		return { text: t('pptx.autosave.saving'), kind: 'saving' };
	}
	if (autosave?.state === 'saved' && typeof autosave.timestamp === 'number') {
		return {
			text: t('pptx.autosave.saved', { time: formatAge(autosave.timestamp, t, now) }),
			kind: 'idle',
		};
	}
	if (autosave?.state === 'error') {
		return { text: t('pptx.autosave.error'), kind: 'error' };
	}
	return {
		text: t(dirty ? 'pptx.statusBar.unsavedChanges' : 'pptx.statusBar.allSaved'),
		kind: 'idle',
	};
}

/** Map a viewer mode string to the pressed view button (unknown modes press none). */
export function statusBarViewMode(
	mode: string | undefined,
	sorterOpen = false,
): StatusBarViewMode | undefined {
	if (mode === 'present') {
		return 'slideShow';
	}
	if (sorterOpen) {
		return 'sorter';
	}
	return mode === 'edit' ? 'normal' : undefined;
}

/** What a host knows about its viewer; {@link statusBarState} derives the bar from it. */
export interface StatusBarHostState {
	slideCount: number;
	/** Zero-based active slide. */
	activeSlideIndex: number;
	dirty: boolean;
	/** The autosave engine's state; undefined when the host has none. */
	autosave?: StatusBarAutosave | undefined;
	/** Undefined when the host has no zoom. */
	zoomPercent?: number | undefined;
	notesExpanded?: boolean | undefined;
	/** The pressed view, see {@link statusBarViewMode}; undefined when none is active. */
	viewMode?: StatusBarViewMode | undefined;
	/**
	 * Controls the host has a handler for; each defaults to true. A control without one is
	 * hidden (`viewModes` drops the whole Normal / Sorter / Slide Show cluster).
	 */
	handles?: Partial<Record<'notes' | 'viewModes' | 'sorter' | 'slideShow', boolean>>;
	/** The host's `hiddenActions`: `zoom`, `notes` and `fullscreen` (the Slide Show button). */
	hiddenActions?: readonly ToolbarActionId[] | undefined;
	translate: StatusBarTranslate;
	/** Clock for the "Saved n minutes ago" text; defaults to now. */
	now?: number;
}

/** The status bar's state for a host: save text, gating and translation in one place. */
export function statusBarState(host: StatusBarHostState): StatusBarViewState {
	const save = resolveStatusBarSave(host.translate, host.autosave, host.dirty, host.now);
	const handles = (id: keyof NonNullable<StatusBarHostState['handles']>) =>
		host.handles?.[id] !== false;
	const hidden = (id: ToolbarActionId) => isActionHidden(id, host.hiddenActions);
	const zoom = hidden('zoom') ? undefined : host.zoomPercent;
	return {
		slideCount: host.slideCount,
		activeSlideIndex: host.activeSlideIndex,
		saveText: save.text,
		saveKind: save.kind,
		...(zoom === undefined ? {} : { zoomPercent: zoom }),
		showNotes: handles('notes') && !hidden('notes'),
		notesExpanded: host.notesExpanded === true,
		showViewModes: handles('viewModes'),
		showSorter: handles('sorter'),
		showSlideShow: handles('slideShow') && !hidden('fullscreen'),
		...(host.viewMode ? { viewMode: host.viewMode } : {}),
		translate: host.translate,
	};
}

/** A host's handler per status bar control; a control without one does nothing. */
export type StatusBarIntentHandlers = Partial<Record<StatusBarControlId, () => void>>;

/** Route the bar's `status-request` event to the host's handler for its control. */
export function dispatchStatusBarIntent(event: Event, handlers: StatusBarIntentHandlers): void {
	const id = (event as CustomEvent<StatusBarIntent | undefined>).detail?.id;
	if (id) {
		handlers[id]?.();
	}
}
