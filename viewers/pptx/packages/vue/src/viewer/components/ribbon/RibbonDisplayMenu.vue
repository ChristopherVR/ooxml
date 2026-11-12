<script setup lang="ts">
import { Check } from 'lucide-vue-next';
/**
 * RibbonDisplayMenu: the Ribbon Display Options menu (Full-screen mode, Show tabs only,
 * Always show Ribbon) under the right end of the tab row. The anchor it sits in has no height,
 * so the tab row's horizontal scroll cannot clip the menu.
 */
import type { RibbonDisplayController, RibbonDisplayState } from 'ooxml-ui/pptx';
import { RIBBON_DISPLAY_OPTIONS_LABEL_KEY } from 'ooxml-ui/pptx';
import { useI18n } from 'vue-i18n';

interface Props {
	state: RibbonDisplayState;
	controller: RibbonDisplayController;
}

const props = defineProps<Props>();
const { t } = useI18n();
</script>

<template>
	<div v-if="props.state.menuOpen" class="pointer-events-none absolute right-1.5 top-0.5 z-40">
		<div
			role="menu"
			data-pptx-chrome="ribbon-display-menu"
			:aria-label="t(RIBBON_DISPLAY_OPTIONS_LABEL_KEY)"
			class="pointer-events-auto w-56 rounded-lg border border-border bg-popover shadow-2xl py-1"
		>
			<button
				v-for="item in props.state.menu"
				:key="item.mode"
				type="button"
				role="menuitemradio"
				:aria-checked="item.checked"
				:data-ribbon-display="item.mode"
				class="flex items-center gap-2 w-full px-3 py-1.5 text-xs text-foreground hover:bg-muted transition-colors"
				@click="props.controller.choose(item.mode)"
			>
				<Check
					class="w-3.5 h-3.5 shrink-0"
					:class="item.checked ? 'text-primary' : 'invisible'"
					aria-hidden="true"
				/>
				{{ t(item.labelKey) }}
			</button>
		</div>
	</div>
</template>
