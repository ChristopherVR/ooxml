<script setup lang="ts">
import { TEXT_EFFECT_SECTIONS } from 'ooxml-ui/pptx';
import type { TextStyle } from 'pptx-viewer-core';
import { useI18n } from 'vue-i18n';

import InspectorFields from './InspectorFields.vue';

/**
 * TextEffectsPanel: text shadow / glow / reflection for the selected text, as
 * React's `TextEffectsPanel` offers them. The fields are the shared
 * `TEXT_EFFECT_SECTIONS` descriptors drawn by the generic {@link InspectorFields}
 * renderer; each change emits an `update` carrying a PARTIAL body-level
 * `TextStyle` patch, which the parent (`TextPanel`) merges onto the current
 * `textStyle`.
 */
const props = withDefaults(
	defineProps<{
		ts: TextStyle | undefined;
		/** Read-only decks show the effects without letting them change. */
		disabled?: boolean;
	}>(),
	{ disabled: false },
);

const emit = defineEmits<{
	update: [patch: Partial<TextStyle>];
}>();

const { t } = useI18n();
</script>

<template>
	<div class="pptx-vue-texteffects mt-2 rounded border border-border bg-card p-2 space-y-2">
		<div class="text-[11px] uppercase tracking-wide text-muted-foreground">
			{{ t('pptx.textEffects.title') }}
		</div>
		<InspectorFields
			:sections="TEXT_EFFECT_SECTIONS"
			:target="props.ts"
			:disabled="props.disabled"
			@patch="emit('update', $event)"
		/>
	</div>
</template>
