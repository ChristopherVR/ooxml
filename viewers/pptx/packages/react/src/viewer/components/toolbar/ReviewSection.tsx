import {
	buildReviewRibbon,
	commentCounts,
	findCommentLocation,
	isReviewCommentCommand,
	runReviewCommentCommand,
	trackActiveComment,
} from 'ooxml-ui/pptx';
import type { ReviewCommentsHost, RibbonControlId } from 'ooxml-ui/pptx';
import React, { useEffect, useState } from 'react';

import { useTranslation } from '../../i18n/use-translation';
import { WebRibbonSection } from './WebRibbonSection';

export interface ReviewSectionProps {
	canEdit: boolean;
	spellCheckEnabled: boolean;
	onSetSpellCheckEnabled: (enabled: boolean) => void;
	onToggleComments?: () => void;
	isCommentsPanelOpen?: boolean;
	slideCommentCount?: number;
	/** Previous / Next / Delete on the deck's comments (absent: those stay disabled). */
	commentsHost?: ReviewCommentsHost;
	onCompare?: () => void;
	onOpenAccessibilityCheck?: () => void;
	onSetLanguage?: () => void;
}

export function ReviewSection(p: ReviewSectionProps): React.ReactElement {
	const { t } = useTranslation();
	const [activeCommentId, setActiveCommentId] = useState<string | null>(null);
	useEffect(() => trackActiveComment(document, setActiveCommentId), []);
	const host = p.commentsHost;
	const slides = host?.slides() ?? [];
	const counts = commentCounts(slides, host?.currentSlide() ?? 0);
	// A comment that was deleted elsewhere is no longer current.
	const hasActiveComment = findCommentLocation(slides, activeCommentId) !== null;
	const groups = buildReviewRibbon(t, {
		editable: p.canEdit,
		spellCheck: p.spellCheckEnabled,
		canAccessibility: Boolean(p.onOpenAccessibilityCheck),
		canLanguage: Boolean(p.onSetLanguage),
		canCompare: Boolean(p.onCompare),
		canComments: Boolean(p.onToggleComments),
		commentsOpen: p.isCommentsPanelOpen,
		commentCount: p.slideCommentCount,
		totalCommentCount: counts.total,
		canNavigateComments: Boolean(host),
		hasActiveComment,
	});
	const request = (id: RibbonControlId) => {
		if (isReviewCommentCommand(id)) {
			const navigation = id === 'review.comments.previous' || id === 'review.comments.next';
			if (host && (navigation || p.canEdit)) {
				const next = runReviewCommentCommand(id, host, activeCommentId);
				if (next !== undefined) {
					setActiveCommentId(next);
				}
			}
			return;
		}
		switch (id) {
			case 'review.proofing.spelling':
				p.onSetSpellCheckEnabled(!p.spellCheckEnabled);
				break;
			case 'review.accessibility.check':
				p.onOpenAccessibilityCheck?.();
				break;
			case 'review.language.language':
				p.onSetLanguage?.();
				break;
			case 'review.compare.compare':
				if (p.canEdit) {
					p.onCompare?.();
				}
				break;
			case 'review.comments.newComment':
			case 'review.comments.showComments':
				p.onToggleComments?.();
				break;
		}
	};
	return <WebRibbonSection groups={groups} onCommand={request} />;
}
