import type { BuildingBlockGroupPorts, PptxBuildingBlockSession } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
import { useEffect, useRef, useState } from 'react';

import { useTranslation } from '../i18n/use-translation';
import { useDialogFlagRequest } from './useDialogFlagRequest';
import type { UsePresentationModeResult } from './usePresentationMode';
import type { ViewerDialogsResult } from './useViewerDialogs';
import type { ViewerState } from './useViewerState';

export type ShowDialogPorts = Pick<
	BuildingBlockGroupPorts,
	'setUpSlideShow' | 'customShows' | 'shortcuts'
>;

export interface UseShowDialogPortsInput {
	state: ViewerState;
	dialogs: ViewerDialogsResult;
	presentation: UsePresentationModeResult;
	slides: PptxSlide[];
	session: Pick<PptxBuildingBlockSession, 'request' | 'announce'>;
}

/**
 * Set Up Slide Show, Custom Shows, Broadcast and the shortcut help: the ports of the shared
 * action groups over this binding's state, plus the effects that turn the toolbar's dialog flags
 * (and the shortcut toggle, `?` and Ctrl+/) into host requests. The operations themselves are the
 * shared groups (`ooxml-ui/pptx`, `render/building-blocks/groups`).
 */
export function useShowDialogPorts(input: UseShowDialogPortsInput): ShowDialogPorts {
	const { state, dialogs, presentation, session } = input;
	const { t } = useTranslation();
	const latest = useRef({ ...input, t });
	latest.current = { ...input, t };
	// Edits made in one tick (rename, then set slides) must build on each other, not on the render.
	const showsRef = useRef(state.customShows);
	showsRef.current = state.customShows;
	const [startRequest, setStartRequest] = useState(0);

	// The toolbar opens its dialogs by flag and nothing renders them here: each becomes a request.
	useDialogFlagRequest(dialogs.isSetUpSlideShowOpen, () => {
		dialogs.setIsSetUpSlideShowOpen(false);
		session.request({ kind: 'setUpSlideShow' });
	});
	useDialogFlagRequest(dialogs.isCustomShowsOpen, () => {
		dialogs.setIsCustomShowsOpen(false);
		session.request({ kind: 'customShows' });
	});
	// Broadcast is unsupported (`actions.broadcast`): the flag is dropped, no request is raised.
	useDialogFlagRequest(dialogs.isBroadcastDialogOpen, () =>
		dialogs.setIsBroadcastDialogOpen(false),
	);

	// The shortcut help is a toggle: every change of its flag (toolbar, `?`, Ctrl+/) is an event.
	const shortcutsOpen = state.isShortcutHelpOpen;
	const firstShortcutsRender = useRef(true);
	useEffect(() => {
		if (firstShortcutsRender.current) {
			firstShortcutsRender.current = false;
			return;
		}
		session.announce({ kind: 'shortcuts', open: shortcutsOpen }, shortcutsOpen);
	}, [shortcutsOpen, session]);

	// Starting a custom show selects it first; the start reads the NEW selection after it commits.
	useEffect(() => {
		if (startRequest > 0) {
			presentation.enterPresentModeFromBeginning();
		}
		// Only a new request starts the show; `presentation` is read at that commit.
		// oxlint-disable-next-line react/exhaustive-deps
	}, [startRequest]);

	// Created once; every member reads the latest render through `latest`.
	const ports = useRef<ShowDialogPorts | null>(null);
	ports.current ??= (() => {
		const now = () => latest.current;
		return {
			setUpSlideShow: {
				presentationProperties: () => now().state.presentationProperties,
				slideCount: () => now().slides.length,
				customShows: () => now().state.customShows,
				save: (settings) => now().dialogs.handleSaveSlideShowSettings(settings),
			},
			customShows: {
				customShows: () => showsRef.current,
				slides: () => now().slides,
				activeId: () => now().state.activeCustomShowId,
				commit: (shows, slides) => {
					showsRef.current = shows;
					now().dialogs.handleCommitCustomShows(shows, slides);
				},
				start: (id) => {
					now().state.setActiveCustomShowId(id);
					setStartRequest((count) => count + 1);
				},
			},
			shortcuts: {
				open: () => now().state.isShortcutHelpOpen,
				setOpen: (open) => now().state.setIsShortcutHelpOpen(open),
				toggle: () => now().state.setIsShortcutHelpOpen((open) => !open),
				translate: (key) => now().t(key),
			},
		};
	})();
	return ports.current;
}
