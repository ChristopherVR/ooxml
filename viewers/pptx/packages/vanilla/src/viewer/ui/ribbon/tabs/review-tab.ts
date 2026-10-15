import {
	buildReviewRibbon,
	commentCounts,
	EMPTY_RESOLVED_CUSTOMIZATION,
	findCommentLocation,
	isDialogAvailable,
	isReviewCommentCommand,
	runReviewCommentCommand,
	trackActiveComment,
} from 'ooxml-ui/pptx';
import type { RibbonCommandRequestEvent } from 'ooxml-ui/pptx';

import type { Translator } from '../../../i18n';
import { createEl } from '../../../render';
import type { RibbonNavHandlers } from '../ribbon-types';

export interface ReviewTab {
	el: HTMLElement;
	setEditable(editable: boolean): void;
	setSpellCheck(enabled: boolean): void;
	/** Release the document listeners that follow the comment the user picks. */
	destroy(): void;
}

/** Shared commands emit intents; the binding retains native panels and editor settings. */
export function createReviewTab(
	doc: Document,
	t: Translator,
	handlers: RibbonNavHandlers,
): ReviewTab {
	const el = createEl(doc, 'div', 'pptxv-ribbon-tab-content');
	const section = doc.createElement('pptx-ui-ribbon-section');
	el.append(section);
	let editable = false;
	let spellCheck = false;
	// The comment Previous / Next / Delete act on; follows the user's own picks in the panel.
	let activeCommentId: string | null = null;
	const refresh = () => {
		const host = handlers.reviewComments;
		const slides = host?.slides() ?? [];
		const counts = commentCounts(slides, host?.currentSlide() ?? 0);
		section.groups = buildReviewRibbon(t, {
			commentCount: counts.slide,
			totalCommentCount: counts.total,
			canNavigateComments: host !== undefined,
			hasActiveComment: findCommentLocation(slides, activeCommentId) !== null,
			editable,
			spellCheck,
			canAccessibility: true,
			canLanguage: true,
			canCompare: true,
			canComments: true,
			languageHidden: !isDialogAvailable(
				handlers.getCustomization?.() ?? EMPTY_RESOLVED_CUSTOMIZATION,
				'options',
			),
		});
	};
	const stopTracking = trackActiveComment(doc, (id) => {
		if (id !== activeCommentId) {
			activeCommentId = id;
			refresh();
		}
	});
	section.addEventListener('command-request', (event) => {
		const id = (event as RibbonCommandRequestEvent).detail.id;
		if (isReviewCommentCommand(id)) {
			const navigation = id === 'review.comments.previous' || id === 'review.comments.next';
			const host = handlers.reviewComments;
			if (host && (navigation || editable)) {
				const next = runReviewCommentCommand(id, host, activeCommentId);
				if (next !== undefined) {
					activeCommentId = next;
					refresh();
				}
			}
			return;
		}
		switch (id) {
			case 'review.proofing.spelling':
				handlers.toggleSpellCheck();
				break;
			case 'review.accessibility.check':
				handlers.openAccessibility();
				break;
			case 'review.language.language':
				handlers.openSettings('general');
				break;
			case 'review.compare.compare':
				if (editable) {
					handlers.openCompare();
				}
				break;
			case 'review.comments.newComment':
			case 'review.comments.showComments':
				handlers.openComments();
				break;
		}
	});
	refresh();
	return {
		el,
		setEditable(value) {
			editable = value;
			refresh();
		},
		setSpellCheck(value) {
			spellCheck = value;
			refresh();
		},
		destroy: stopTracking,
	};
}
