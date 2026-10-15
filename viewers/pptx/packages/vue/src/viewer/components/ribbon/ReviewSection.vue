<script setup lang="ts">
import {
	buildReviewRibbon,
	commentCounts,
	findCommentLocation,
	isReviewCommentCommand,
	runReviewCommentCommand,
	trackActiveComment,
} from 'ooxml-ui/pptx';
import type { RibbonCommandRequestEvent, ReviewCommentsHost } from 'ooxml-ui/pptx';
import { computed, onBeforeUnmount, onMounted, ref } from 'vue';
import { useI18n } from 'vue-i18n';

interface Props {
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

const props = defineProps<Props>();

const { t } = useI18n();
const activeCommentId = ref<string | null>(null);
let stopTracking: (() => void) | undefined;
onMounted(() => {
	stopTracking = trackActiveComment(document, (id) => {
		activeCommentId.value = id;
	});
});
onBeforeUnmount(() => stopTracking?.());
const groups = computed(() => {
	const host = props.commentsHost;
	const slides = host?.slides() ?? [];
	const counts = commentCounts(slides, host?.currentSlide() ?? 0);
	return buildReviewRibbon(t, {
		editable: props.canEdit,
		spellCheck: props.spellCheckEnabled,
		canAccessibility: Boolean(props.onOpenAccessibilityCheck),
		canLanguage: Boolean(props.onSetLanguage),
		canCompare: Boolean(props.onCompare),
		canComments: Boolean(props.onToggleComments),
		commentsOpen: props.isCommentsPanelOpen,
		commentCount: props.slideCommentCount,
		totalCommentCount: counts.total,
		canNavigateComments: Boolean(host),
		// A comment deleted elsewhere is no longer current.
		hasActiveComment: findCommentLocation(slides, activeCommentId.value) !== null,
	});
});
function request(event: Event): void {
	const id = (event as RibbonCommandRequestEvent).detail.id;
	if (isReviewCommentCommand(id)) {
		const navigation = id === 'review.comments.previous' || id === 'review.comments.next';
		if (props.commentsHost && (navigation || props.canEdit)) {
			const next = runReviewCommentCommand(id, props.commentsHost, activeCommentId.value);
			if (next !== undefined) {
				activeCommentId.value = next;
			}
		}
		return;
	}
	switch (id) {
		case 'review.proofing.spelling':
			props.onSetSpellCheckEnabled(!props.spellCheckEnabled);
			break;
		case 'review.accessibility.check':
			props.onOpenAccessibilityCheck?.();
			break;
		case 'review.language.language':
			props.onSetLanguage?.();
			break;
		case 'review.compare.compare':
			if (props.canEdit) {
				props.onCompare?.();
			}
			break;
		case 'review.comments.newComment':
		case 'review.comments.showComments':
			props.onToggleComments?.();
			break;
	}
}
</script>
<template><pptx-ui-ribbon-section :groups.prop="groups" @command-request="request" /></template>
