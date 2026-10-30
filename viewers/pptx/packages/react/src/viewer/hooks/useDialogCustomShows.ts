import { generateCustomShowId } from 'ooxml-ui/pptx';
import type { PptxCustomShow, PptxSlide } from 'pptx-viewer-core';
/**
 * Custom shows dialog handlers: extracted from useViewerDialogs.
 */
import { useMemo } from 'react';

import { safePrompt, safeConfirm } from '../utils/dom-helpers';
import type { EditorHistoryResult } from './useEditorHistory';

export interface UseDialogCustomShowsInput {
	activeSlide: PptxSlide | undefined;
	customShows: Array<{ id: string; name: string; slideRIds: string[] }>;
	activeCustomShowId: string | null;
	setCustomShows: React.Dispatch<
		React.SetStateAction<Array<{ id: string; name: string; slideRIds: string[] }>>
	>;
	setActiveCustomShowId: React.Dispatch<React.SetStateAction<string | null>>;
	history: EditorHistoryResult;
	/** Replaces the deck's slides (used when removing a show clears jumps into it). */
	setSlides?: React.Dispatch<React.SetStateAction<PptxSlide[]>>;
}

export interface UseDialogCustomShowsResult {
	/** OK in the Custom Shows dialog: store the shows, clear jumps into removed ones. */
	handleCommitCustomShows: (
		shows: PptxCustomShow[],
		clearedSlides: PptxSlide[] | undefined,
	) => void;
	handleCreateCustomShow: () => void;
	handleRenameActiveCustomShow: () => void;
	handleDeleteActiveCustomShow: () => void;
	handleToggleCurrentSlideInActiveShow: () => void;
	isCurrentSlideInActiveShow: boolean;
}

export function useDialogCustomShows(input: UseDialogCustomShowsInput): UseDialogCustomShowsResult {
	const {
		activeSlide,
		customShows,
		activeCustomShowId,
		setCustomShows,
		setActiveCustomShowId,
		history,
		setSlides,
	} = input;

	const handleCommitCustomShows = (
		shows: PptxCustomShow[],
		clearedSlides: PptxSlide[] | undefined,
	) => {
		setCustomShows(shows);
		if (clearedSlides) {
			setSlides?.(clearedSlides);
		}
		setActiveCustomShowId((current) =>
			current !== null && shows.some((show) => show.id === current) ? current : null,
		);
		history.markDirty();
	};

	const handleCreateCustomShow = () => {
		const name = safePrompt('Custom show name', 'Custom Show');
		const safeName = name?.trim() || `Custom Show ${customShows.length + 1}`;
		// `p:custShow/@id` is an xsd:unsignedInt, so ids are numbers, never random strings.
		const id = generateCustomShowId(customShows);
		const initialSlideRIds =
			activeSlide?.rId && activeSlide.rId.length > 0 ? [activeSlide.rId] : [];
		setCustomShows((prev) => [...prev, { id, name: safeName, slideRIds: initialSlideRIds }]);
		setActiveCustomShowId(id);
		history.markDirty();
	};

	const handleRenameActiveCustomShow = () => {
		if (!activeCustomShowId) {
			return;
		}
		const show = customShows.find((s) => s.id === activeCustomShowId);
		if (!show) {
			return;
		}
		const nextName = safePrompt('Rename custom show', show.name)?.trim();
		if (!nextName) {
			return;
		}
		setCustomShows((prev) => prev.map((s) => (s.id === show.id ? { ...s, name: nextName } : s)));
		history.markDirty();
	};

	const handleDeleteActiveCustomShow = () => {
		const showId = activeCustomShowId;
		if (!showId) {
			return;
		}
		const show = customShows.find((s) => s.id === showId);
		if (!show) {
			return;
		}
		if (!safeConfirm(`Delete custom show "${show.name}"?`)) {
			return;
		}
		setCustomShows((prev) => prev.filter((s) => s.id !== showId));
		setActiveCustomShowId(null);
		history.markDirty();
	};

	const handleToggleCurrentSlideInActiveShow = () => {
		const showId = activeCustomShowId;
		const slideRId = activeSlide?.rId;
		if (!showId || !slideRId) {
			return;
		}
		setCustomShows((prev) =>
			prev.map((s) => {
				if (s.id !== showId) {
					return s;
				}
				const hasSlide = s.slideRIds.includes(slideRId);
				return {
					...s,
					slideRIds: hasSlide
						? s.slideRIds.filter((rid) => rid !== slideRId)
						: [...s.slideRIds, slideRId],
				};
			}),
		);
		history.markDirty();
	};

	const isCurrentSlideInActiveShow = useMemo(() => {
		if (!activeCustomShowId || !activeSlide?.rId) {
			return false;
		}
		const show = customShows.find((s) => s.id === activeCustomShowId);
		return show ? show.slideRIds.includes(activeSlide.rId) : false;
	}, [activeCustomShowId, activeSlide?.rId, customShows]);

	return {
		handleCommitCustomShows,
		handleCreateCustomShow,
		handleRenameActiveCustomShow,
		handleDeleteActiveCustomShow,
		handleToggleCurrentSlideInActiveShow,
		isCurrentSlideInActiveShow,
	};
}
