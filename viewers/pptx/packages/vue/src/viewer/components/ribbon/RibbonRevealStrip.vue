<script setup lang="ts">
/**
 * RibbonRevealStrip: the slim strip that brings the ribbon back over the slide in full-screen
 * mode. It stays in the layout while the ribbon is peeked, so nothing jumps when it appears.
 */
import type { RibbonDisplayController } from 'ooxml-ui/pptx';
import { RIBBON_DISPLAY_REVEAL_LABEL_KEY } from 'ooxml-ui/pptx';
import { useI18n } from 'vue-i18n';

interface Props {
	controller: RibbonDisplayController;
}

const props = defineProps<Props>();
const { t } = useI18n();
</script>

<template>
	<button
		type="button"
		data-pptx-chrome="ribbon-reveal"
		:title="t(RIBBON_DISPLAY_REVEAL_LABEL_KEY)"
		:aria-label="t(RIBBON_DISPLAY_REVEAL_LABEL_KEY)"
		class="flex h-3.5 w-full shrink-0 items-center justify-center border-b border-border bg-secondary/50 text-[10px] leading-none text-muted-foreground transition-colors hover:text-foreground"
		@click="props.controller.reveal()"
	>
		···
	</button>
</template>
