import {
	BUILDING_BLOCK_OK,
	applyCustomShowEdit,
	buildingBlockFailure,
	describeBroadcastSupport,
	describeCustomShows,
	describeShortcutCatalogue,
	describeShowDialogActions,
	describeSlideShowSettings,
	resolveCustomShowStart,
	resolveSlideShowSettings,
} from 'ooxml-ui/pptx';
import type {
	BroadcastSupport,
	BuildingBlockActionState,
	BuildingBlockAvailabilityInput,
	BuildingBlockResult,
	CustomShowEdit,
	CustomShowSlideChoice,
	CustomShowSummary,
	ShortcutCatalogueEntry,
	ShowDialogActionId,
	ShowDialogRequest,
	SlideShowSettingsPatch,
	SlideShowSettingsView,
} from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';

import { useTranslation } from '../i18n/use-translation';
import type { EditorHistoryResult } from './useEditorHistory';
import type { UsePresentationModeResult } from './usePresentationMode';
import type { ViewerDialogsResult } from './useViewerDialogs';
import type { ViewerState } from './useViewerState';

/** Set Up Slide Show, as a host's own dialog sees it. */
export interface BuildingBlockSetUpSlideShow {
	/** The current settings with every default resolved. */
	current: SlideShowSettingsView;
	/** Store the changed fields (the dialog's OK). Validated; the deck is marked dirty. */
	apply: (patch: SlideShowSettingsPatch) => BuildingBlockResult;
}

/** Custom Shows, as a host's own dialog sees it. Every edit is stored at once. */
export interface BuildingBlockCustomShows {
	shows: CustomShowSummary[];
	/** The deck's slides a show may list, in deck order. */
	slides: CustomShowSlideChoice[];
	/** The show Set Up and the Slide Show tab currently use, or null. */
	activeId: string | null;
	/** Create a show (the active slide when `slideRIds` is omitted is not assumed). */
	create: (name: string, slideRIds?: readonly string[]) => BuildingBlockResult;
	rename: (id: string, name: string) => BuildingBlockResult;
	/** Replace the show's slide list (relationship ids, in show order). */
	setSlides: (id: string, slideRIds: readonly string[]) => BuildingBlockResult;
	copy: (id: string, name?: string) => BuildingBlockResult;
	/** Delete a show; jumps into it are cleared. */
	remove: (id: string) => BuildingBlockResult;
	/** Make `id` the active show and start the slide show from its first slide. */
	start: (id: string) => BuildingBlockResult;
}

export interface BuildingBlockShortcuts {
	/** The shortcut help is open (also toggled by `?` and Ctrl+/). The host renders it. */
	open: boolean;
	setOpen: (open: boolean) => void;
	toggle: () => void;
	/** The keymap as data, in the viewer's current language. */
	catalogue: ShortcutCatalogueEntry[];
}

export interface ShowDialogActions {
	setUpSlideShow: BuildingBlockSetUpSlideShow;
	customShows: BuildingBlockCustomShows;
	/** Broadcast cannot be completed by the building blocks; see {@link BroadcastSupport}. */
	broadcast: BroadcastSupport;
	shortcuts: BuildingBlockShortcuts;
}

export interface UseShowDialogActionsInput {
	state: ViewerState;
	dialogs: ViewerDialogsResult;
	history: EditorHistoryResult;
	presentation: UsePresentationModeResult;
	slides: PptxSlide[];
	/** The inputs `describeBuildingBlockActions` gets, so the edit gate is the same. */
	availabilityInput: BuildingBlockAvailabilityInput;
	onRequest?: (request: ShowDialogRequest) => void;
}

export interface ShowDialogActionsResult {
	actions: ShowDialogActions;
	availability: Record<ShowDialogActionId, BuildingBlockActionState>;
	pending: ShowDialogRequest | null;
	cancel: () => void;
}

/**
 * The host-owned slide show dialogs of the building-block API. The toolbar's Set Up Slide Show,
 * Custom Shows and shortcut controls (and Ctrl+/) raise a request; the operations run the
 * viewer's own settings and custom show code, re-check the edit gate when they execute, and mark
 * the deck dirty as the bundled dialogs do.
 */
export function useViewerBuildingBlocksShowDialogActions(
	input: UseShowDialogActionsInput,
): ShowDialogActionsResult {
	const { state, dialogs, presentation } = input;
	const { t } = useTranslation();
	const [pending, setPending] = useState<ShowDialogRequest | null>(null);
	const [startRequest, setStartRequest] = useState(0);

	const availability = describeShowDialogActions(input.availabilityInput);
	const latest = useRef({ ...input, availability });
	latest.current = { ...input, availability };
	const refuse = (id: ShowDialogActionId): BuildingBlockResult | undefined => {
		const { available, reason } = latest.current.availability[id];
		return available ? undefined : buildingBlockFailure(reason ?? 'not-ready');
	};
	// Edits made in one tick (rename, then set slides) must build on each other, not on the render.
	const showsRef = useRef(state.customShows);
	showsRef.current = state.customShows;
	const context = () => ({
		slideCount: latest.current.slides.length,
		customShows: latest.current.state.customShows,
	});

	const raise = useCallback((request: ShowDialogRequest, id: ShowDialogActionId) => {
		if (!latest.current.availability[id].available) {
			return;
		}
		setPending(request);
		latest.current.onRequest?.(request);
	}, []);

	// The toolbar opens its dialogs by flag and nothing renders them here: each becomes a request.
	const { isSetUpSlideShowOpen, isCustomShowsOpen, isBroadcastDialogOpen } = dialogs;
	useEffect(() => {
		if (isSetUpSlideShowOpen) {
			dialogs.setIsSetUpSlideShowOpen(false);
			raise({ kind: 'setUpSlideShow' }, 'setUpSlideShow');
		}
	}, [isSetUpSlideShowOpen, dialogs, raise]);
	useEffect(() => {
		if (isCustomShowsOpen) {
			dialogs.setIsCustomShowsOpen(false);
			raise({ kind: 'customShows' }, 'customShows');
		}
	}, [isCustomShowsOpen, dialogs, raise]);
	useEffect(() => {
		// Broadcast is unsupported (`actions.broadcast`): the flag is dropped, no request is raised.
		if (isBroadcastDialogOpen) {
			dialogs.setIsBroadcastDialogOpen(false);
		}
	}, [isBroadcastDialogOpen, dialogs]);

	// The shortcut help is a toggle: every change of its flag (toolbar, `?`, Ctrl+/) is an event.
	const shortcutsOpen = state.isShortcutHelpOpen;
	const firstShortcutsRender = useRef(true);
	useEffect(() => {
		if (firstShortcutsRender.current) {
			firstShortcutsRender.current = false;
			return;
		}
		const request: ShowDialogRequest = { kind: 'shortcuts', open: shortcutsOpen };
		setPending(
			shortcutsOpen ? request : (current) => (current?.kind === 'shortcuts' ? null : current),
		);
		latest.current.onRequest?.(request);
	}, [shortcutsOpen]);

	// Starting a custom show selects it first; the start reads the NEW selection after it commits.
	useEffect(() => {
		if (startRequest > 0) {
			presentation.enterPresentModeFromBeginning();
		}
		// Only a new request starts the show; `presentation` is read at that commit.
		// oxlint-disable-next-line react/exhaustive-deps
	}, [startRequest]);

	const complete = (kind: ShowDialogRequest['kind'], result: BuildingBlockResult) => {
		if (result.ok) {
			setPending((current) => (current?.kind === kind ? null : current));
		}
		return result;
	};

	const setUpSlideShow: BuildingBlockSetUpSlideShow = {
		current: describeSlideShowSettings(state.presentationProperties, {
			slideCount: input.slides.length,
			customShows: state.customShows,
		}),
		apply: (patch) => {
			const denied = refuse('setUpSlideShow');
			if (denied) {
				return denied;
			}
			const resolved = resolveSlideShowSettings(
				latest.current.state.presentationProperties,
				patch,
				context(),
			);
			if (!resolved.ok) {
				return buildingBlockFailure(resolved.code);
			}
			latest.current.dialogs.handleSaveSlideShowSettings(resolved.settings);
			return complete('setUpSlideShow', BUILDING_BLOCK_OK);
		},
	};

	const edit = (change: CustomShowEdit): BuildingBlockResult => {
		const denied = refuse('customShows');
		if (denied) {
			return denied;
		}
		const now = latest.current;
		const outcome = applyCustomShowEdit(showsRef.current, now.slides, change);
		if (!outcome.ok) {
			return buildingBlockFailure(outcome.code);
		}
		showsRef.current = outcome.shows;
		now.dialogs.handleCommitCustomShows(outcome.shows, outcome.slides);
		return BUILDING_BLOCK_OK;
	};
	const described = describeCustomShows(state.customShows, input.slides);
	const customShows: BuildingBlockCustomShows = {
		...described,
		activeId: state.activeCustomShowId,
		create: (name, slideRIds) =>
			edit({ kind: 'create', name, ...(slideRIds ? { slideRIds } : {}) }),
		rename: (id, name) => edit({ kind: 'rename', id, name }),
		setSlides: (id, slideRIds) => edit({ kind: 'setSlides', id, slideRIds }),
		copy: (id, name) => edit({ kind: 'copy', id, ...(name === undefined ? {} : { name }) }),
		remove: (id) => edit({ kind: 'remove', id }),
		start: (id) => {
			const now = latest.current;
			if (!now.availabilityInput.ready) {
				return buildingBlockFailure('not-ready');
			}
			const target = resolveCustomShowStart(now.state.customShows, now.slides, id);
			if (!target.ok) {
				return buildingBlockFailure(target.code);
			}
			now.state.setActiveCustomShowId(id);
			setStartRequest((count) => count + 1);
			return complete('customShows', BUILDING_BLOCK_OK);
		},
	};

	const setShortcutsOpen = useCallback(
		(open: boolean) => latest.current.state.setIsShortcutHelpOpen(open),
		[],
	);
	const catalogue = useMemo(() => describeShortcutCatalogue((key) => t(key)), [t]);
	const shortcuts: BuildingBlockShortcuts = {
		open: shortcutsOpen,
		setOpen: setShortcutsOpen,
		toggle: () => latest.current.state.setIsShortcutHelpOpen((open) => !open),
		catalogue,
	};

	return {
		actions: { setUpSlideShow, customShows, broadcast: describeBroadcastSupport(), shortcuts },
		availability,
		pending,
		cancel: () => {
			setPending(null);
			latest.current.state.setIsShortcutHelpOpen(false);
		},
	};
}
