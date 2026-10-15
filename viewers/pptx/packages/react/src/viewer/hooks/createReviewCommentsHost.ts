import type { ReviewCommentsHost } from 'ooxml-ui/pptx';
import type { PptxSlide } from 'pptx-viewer-core';
import type React from 'react';

/** What Review > Comments Previous / Next / Delete need from the React viewer state. */
export interface ReviewCommentsHostInput {
	state: {
		slides: PptxSlide[];
		isInspectorPaneOpen: boolean;
		setIsInspectorPaneOpen: React.Dispatch<React.SetStateAction<boolean>>;
		setSidebarPanelMode: React.Dispatch<React.SetStateAction<string>>;
		setActiveSlideIndex: React.Dispatch<React.SetStateAction<number>>;
	};
	activeSlideIndex: number;
	updateSlides: (updater: (slides: PptxSlide[]) => PptxSlide[]) => void;
	markDirty: () => void;
}

/** The host both toolbar assemblies (`buildToolbarProps`, `ViewerToolbarSection`) hand the ribbon. */
export function createReviewCommentsHost(input: ReviewCommentsHostInput): ReviewCommentsHost {
	const { state: s, activeSlideIndex, updateSlides, markDirty } = input;
	return {
		slides: () => s.slides,
		currentSlide: () => activeSlideIndex,
		goToSlide: (index) => s.setActiveSlideIndex(index),
		ensureCommentsOpen: () => {
			s.setSidebarPanelMode('comments');
			if (!s.isInspectorPaneOpen) {
				s.setIsInspectorPaneOpen(true);
			}
		},
		applyComments: (changes) => {
			updateSlides((slides) =>
				slides.map((slide, index) => {
					const change = changes.find((candidate) => candidate.slideIndex === index);
					return change ? { ...slide, comments: change.comments } : slide;
				}),
			);
			markDirty();
		},
	};
}
