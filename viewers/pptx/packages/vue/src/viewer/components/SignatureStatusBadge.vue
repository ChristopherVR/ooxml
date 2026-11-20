<script setup lang="ts">
import type { SignatureStatusView } from 'ooxml-ui/pptx';

import { SIGNATURE_BADGE_TONE, SIGNATURE_ICONS } from './signature-look';

/**
 * SignatureStatusBadge: a small chrome badge showing the deck's signature state
 * ("Signed" only when the signatures verify). Vue port of the React
 * `SignatureStatusBadge.tsx`. Renders nothing when the deck carries no
 * signatures; emits `click` to open the signatures panel/dialog.
 */
const props = defineProps<{
	view: SignatureStatusView;
}>();

const emit = defineEmits<{
	click: [];
}>();
</script>

<template>
	<button
		v-if="props.view.present"
		type="button"
		class="inline-flex items-center gap-1 rounded-md border px-2 py-0.5 text-[11px] font-medium transition-colors"
		:class="SIGNATURE_BADGE_TONE[props.view.tone]"
		:data-pptx-signature-state="props.view.state"
		:title="props.view.description"
		:aria-label="props.view.description"
		@click="emit('click')"
	>
		<component :is="SIGNATURE_ICONS[props.view.icon]" class="h-3.5 w-3.5" aria-hidden="true" />
		<span>{{ props.view.label }}</span>
	</button>
</template>
