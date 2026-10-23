<script setup lang="ts" generic="T, P">
import {
	createInspectorColorFields,
	handleLiveColorEvent,
	inspectorFieldPatch,
	inspectorTogglePatch,
	resolveInspectorSections,
} from 'ooxml-ui/pptx';
import type { InspectorFieldValue, InspectorSectionDescriptor } from 'ooxml-ui/pptx';
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';

import { injectRecentColors } from '../../composables/recent-colors-context';

/**
 * InspectorFields: the Vue binding's one generic field renderer. It draws any
 * list of shared inspector section descriptors (a toggle per section, then a
 * two-column grid of its fields); values, bounds, labels and patches all come
 * from the descriptors, so this SFC only maps them onto the template. A colour
 * field previews while its native picker is dragged (`preview`, no undo step)
 * and commits once when the pick settles (`patch`, plus the "Recent colours"
 * list), through the shared `createInspectorColorFields`. `targetKey` names the
 * element `target` belongs to, so a field tells a new selection from a
 * re-render of its own drag.
 */
const props = withDefaults(
	defineProps<{
		sections: readonly InspectorSectionDescriptor<T, P>[];
		target: T;
		/** The identity of the element `target` belongs to (its id), for the colour fields. */
		targetKey?: string;
		disabled?: boolean;
	}>(),
	{ disabled: false, targetKey: '' },
);

const emit = defineEmits<{
	patch: [patch: P];
	/**
	 * A colour drag's live patch for the element `target` names (the current one,
	 * or one the drag has left): apply it without an undo step.
	 */
	preview: [patch: P, target: string];
}>();

const { t } = useI18n();
const recentColors = injectRecentColors();

const views = computed(() => resolveInspectorSections(props.sections, props.target));

function onField(sectionId: string, fieldId: string, raw: InspectorFieldValue): void {
	const patch = inspectorFieldPatch(props.sections, sectionId, fieldId, props.target, raw);
	if (patch !== undefined) {
		emit('patch', patch);
	}
}

function onToggle(sectionId: string, on: boolean): void {
	const patch = inspectorTogglePatch(props.sections, sectionId, props.target, on);
	if (patch !== undefined) {
		emit('patch', patch);
	}
}

/** One live colour controller per colour field (shared `createInspectorColorFields`). */
const colorField = createInspectorColorFields<T, P>({
	sections: () => props.sections,
	target: () => props.target,
	targetKey: () => props.targetKey,
	preview: (patch, target) => emit('preview', patch, target),
	commit: (patch) => emit('patch', patch),
	remember: (hex) => recentColors?.push(hex),
});

function onColor(sectionId: string, fieldId: string, event: Event): void {
	handleLiveColorEvent(colorField(sectionId, fieldId), event);
}

const valueOf = (event: Event): string => (event.target as HTMLInputElement).value;
const checkedOf = (event: Event): boolean => (event.target as HTMLInputElement).checked;
</script>

<template>
	<div
		v-for="section in views"
		:key="section.id"
		class="pptx-vue-fields-section flex flex-col gap-2 border-b border-border pb-2 last:border-b-0 last:pb-0"
		:data-section="section.id"
	>
		<label
			v-if="section.toggleable"
			class="pptx-vue-fields-check inline-flex items-center gap-2 text-foreground"
		>
			<pptx-ui-checkbox
				:data-testid="section.toggleTestId"
				:checked="section.on"
				:disabled="props.disabled"
				@change="onToggle(section.id, checkedOf($event))"
			/>
			{{ t(section.labelKey ?? '') }}
		</label>
		<span v-else-if="section.labelKey" class="pptx-vue-fields-title text-foreground">{{
			t(section.labelKey)
		}}</span>
		<div v-if="section.fields.length" class="pptx-vue-fields-grid grid grid-cols-2 gap-2">
			<template v-for="field in section.fields" :key="field.id">
				<label
					v-if="field.control === 'checkbox'"
					class="pptx-vue-fields-check col-span-2 inline-flex items-center gap-2 text-foreground"
				>
					<pptx-ui-checkbox
						:data-testid="field.testId"
						:checked="Boolean(field.value)"
						:disabled="props.disabled"
						@change="onField(section.id, field.id, checkedOf($event))"
					/>
					{{ t(field.labelKey) }}
				</label>
				<label
					v-else
					class="pptx-vue-fields-field flex flex-col gap-1"
					:class="{ 'col-span-2': field.wide }"
				>
					<span class="pptx-vue-fields-label text-muted-foreground">{{ t(field.labelKey) }}</span>
					<pptx-ui-select
						v-if="field.control === 'select'"
						class="pptx-vue-fields-select w-full rounded border border-border bg-muted px-2 py-1 max-md:min-h-[44px]"
						:data-testid="field.testId"
						:aria-label="t(field.labelKey)"
						:value="String(field.value)"
						:disabled="props.disabled"
						@change="onField(section.id, field.id, valueOf($event))"
					>
						<option v-for="option in field.options" :key="option.value" :value="option.value">
							{{ t(option.labelKey) }}
						</option>
					</pptx-ui-select>
					<input
						v-else-if="field.control === 'color'"
						type="color"
						class="pptx-vue-fields-color h-7 w-full rounded border border-border bg-muted p-0"
						:data-testid="field.testId"
						:aria-label="t(field.labelKey)"
						:value="String(field.value)"
						:disabled="props.disabled"
						@input="onColor(section.id, field.id, $event)"
						@change="onColor(section.id, field.id, $event)"
						@blur="onColor(section.id, field.id, $event)"
					/>
					<input
						v-else
						:type="field.control"
						:class="
							field.control === 'range'
								? 'pptx-vue-fields-range w-full accent-primary'
								: 'pptx-vue-fields-input rounded border border-border bg-muted px-2 py-1 max-md:min-h-[44px]'
						"
						:data-testid="field.testId"
						:aria-label="t(field.labelKey)"
						:min="field.min"
						:max="field.max"
						:step="field.step"
						:value="field.value"
						:disabled="props.disabled"
						@input="onField(section.id, field.id, valueOf($event))"
					/>
				</label>
			</template>
		</div>
	</div>
</template>
