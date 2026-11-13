<script setup lang="ts">
/**
 * SlideTextRunContent - a run's content, inside the Find bar's
 * `<mark data-pptx-find-match>` when shared `highlightFindMatches` marked the
 * run as (part of) a match. The text itself is `SlideTextRunPieces`.
 */
import type { ParagraphRun } from 'ooxml-ui/pptx';
import { findMatchMarkStyle } from 'ooxml-ui/pptx';

import SlideTextRunPieces from './SlideTextRunPieces.vue';

defineProps<{ run: ParagraphRun }>();
</script>

<template>
	<mark
		v-if="run.findMatch"
		:data-pptx-find-match="run.findMatch"
		:style="findMatchMarkStyle(run.findMatch)"
		><SlideTextRunPieces :run="run"
	/></mark>
	<SlideTextRunPieces v-else :run="run" />
</template>
