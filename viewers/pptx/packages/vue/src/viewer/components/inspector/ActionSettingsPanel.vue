<script setup lang="ts">
import { ACTION_SETTINGS_SECTIONS, actionSettingsTargetOf } from 'ooxml-ui/pptx';
import type { ActionSettingsChange } from 'ooxml-ui/pptx';
/**
 * ActionSettingsPanel: PowerPoint's Insert > Action dialog as an inspector
 * card. Each trigger (`actionClick` / `actionHover`) is the shared
 * `ACTION_SETTINGS_SECTIONS` drawn by the generic renderer over its context
 * (`actionSettingsTargetOf`: its action, the pending pick and the deck's
 * slides and custom shows); a change holds its type as pending and, once it
 * can carry its target, commits through core's `elementActionToPptxAction`.
 */
import type { ElementActionType, PptxCustomShow, PptxElement } from 'pptx-viewer-core';
import { elementActionToPptxAction } from 'pptx-viewer-core';
import { ref, watch } from 'vue';
import { useI18n } from 'vue-i18n';

import InspectorFields from './InspectorFields.vue';

type Trigger = 'click' | 'hover';

const props = withDefaults(
	defineProps<{
		element: PptxElement;
		slideCount?: number;
		canEdit?: boolean;
		/** Named custom shows, for the "Custom show" target picker. */
		customShows?: readonly PptxCustomShow[];
	}>(),
	{ slideCount: 0, canEdit: true },
);
const emit = defineEmits<{ update: [patch: Partial<PptxElement>] }>();
const { t } = useI18n();

const triggers: readonly Trigger[] = ['click', 'hover'];

/** The type picked but not committed, per trigger; it belongs to the element it was made on. */
const pending = ref<Partial<Record<Trigger, ElementActionType>>>({});
watch(
	() => props.element.id,
	() => {
		pending.value = {};
	},
);

function targetOf(trigger: Trigger) {
	return actionSettingsTargetOf(
		props.element,
		trigger,
		pending.value[trigger],
		props.slideCount,
		(props.customShows ?? []).map((show) => ({ id: show.id, name: show.name })),
	);
}

function apply(trigger: Trigger, change: ActionSettingsChange): void {
	pending.value = { ...pending.value, [trigger]: change.pendingType };
	if (change.commit) {
		const action = elementActionToPptxAction(change.commit);
		emit(
			'update',
			(trigger === 'click'
				? { actionClick: action }
				: { actionHover: action }) as Partial<PptxElement>,
		);
	}
}
</script>

<template>
	<div class="pptx-vue-action-settings flex flex-col gap-3 text-[11px]" data-pptx-action-settings>
		<div
			v-for="trigger in triggers"
			:key="trigger"
			class="flex flex-col gap-1.5"
			:data-pptx-action-trigger="trigger"
		>
			<span class="font-medium text-muted-foreground">{{
				trigger === 'click' ? t('pptx.action.onClick') : t('pptx.action.onHover')
			}}</span>
			<InspectorFields
				:sections="ACTION_SETTINGS_SECTIONS"
				:target="targetOf(trigger)"
				:disabled="!props.canEdit"
				@patch="apply(trigger, $event)"
			/>
		</div>
	</div>
</template>
