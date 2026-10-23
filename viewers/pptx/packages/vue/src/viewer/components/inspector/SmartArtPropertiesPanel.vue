<script setup lang="ts">
import {
	schemaLabel,
	SMARTART_COLOR_SCHEME_LABEL_KEYS,
	SMARTART_STYLE_LABEL_KEYS,
} from 'ooxml-ui/pptx';
import type { SmartArtTextPaneChangeEvent, SmartArtTextPaneViewState } from 'ooxml-ui/pptx';
import type {
	PptxElement,
	PptxSmartArtData,
	SmartArtColorScheme,
	SmartArtLayoutType,
	SmartArtStyle,
} from 'pptx-viewer-core';
import { computed } from 'vue';
import { useI18n } from 'vue-i18n';

import { injectRecentColors } from '../../composables/recent-colors-context';
import {
	SMARTART_COLOR_SCHEMES,
	SMARTART_STYLE_OPTIONS,
	useSmartArtEditing,
} from '../../composables/useSmartArtEditing';
import SmartArtLayoutSwitcher from './SmartArtLayoutSwitcher.vue';

/**
 * SmartArtPropertiesPanel: inspector panel for `smartArt` elements.
 *
 * Vue port of the React `SmartArtPropertiesPanel.tsx`: a layout switcher,
 * colour scheme and style, and the text pane, which is the shared
 * `pptx-ui-smartart-text-pane` element (outline with levels, keyboard structure
 * edits, move, sub-item, remove, per-node style, layout bounds and the non-tree
 * connections note). This SFC feeds it the data and commits what it emits
 * through `useSmartArtEditing`.
 *
 * Follows the uniform inspector-panel contract: `props { element }`, emits
 * `update` with a shallow `Partial<PptxElement>` patch the host merges via
 * `useEditorOperations.updateElement` (history-tracked, so undo/redo works).
 */
const props = withDefaults(
	defineProps<{
		element: PptxElement;
		canEdit?: boolean;
	}>(),
	{ canEdit: true },
);

const emit = defineEmits<{ update: [patch: Partial<PptxElement>] }>();

const { t } = useI18n();
const recentColors = injectRecentColors();

const editable = computed(() => props.canEdit !== false);

const smartArtData = computed<PptxSmartArtData>(() => {
	if (props.element.type === 'smartArt' && props.element.smartArtData) {
		return props.element.smartArtData;
	}
	return { nodes: [] } as PptxSmartArtData;
});

const isSmartArt = computed(
	() => props.element.type === 'smartArt' && Boolean(props.element.smartArtData),
);

const box = computed(() => ({ width: props.element.width, height: props.element.height }));

const editing = useSmartArtEditing({
	smartArtData,
	apply: (patch) => emit('update', patch),
	elementId: props.element.id,
	box,
});

const colorSchemes = SMARTART_COLOR_SCHEMES;
const styleOptions = SMARTART_STYLE_OPTIONS;

/**
 * Spell the `dgm:colorsDef` / `dgm:styleDef` families for display.
 *
 * Both controls keep their existing value sets and still emit the raw wire
 * tokens; only the visible text changes, so nothing is added or dropped.
 * `t` is an overloaded generic, hence the narrowing lambda.
 */
function colorSchemeLabel(scheme: SmartArtColorScheme): string {
	return schemaLabel(SMARTART_COLOR_SCHEME_LABEL_KEYS, scheme, (key: string) => t(key));
}

function styleLabel(style: SmartArtStyle): string {
	return schemaLabel(SMARTART_STYLE_LABEL_KEYS, style, (key: string) => t(key));
}

const paneState = computed<SmartArtTextPaneViewState>(() => ({
	smartArtData: isSmartArt.value ? smartArtData.value : undefined,
	canEdit: editable.value,
	translate: (key, params) => (params ? t(key, { ...params }) : t(key)),
}));

function onPaneChange(event: Event): void {
	const { smartArtData: next, recentColor } = (event as SmartArtTextPaneChangeEvent).detail;
	editing.commit(next);
	if (recentColor) {
		recentColors?.push(recentColor);
	}
}

function onColorSchemeChange(event: Event): void {
	editing.setColorScheme((event.target as HTMLSelectElement).value as SmartArtColorScheme);
}

function onSwitchLayout(layout: SmartArtLayoutType): void {
	editing.switchLayout(layout);
}

function onSetStyle(value: SmartArtStyle): void {
	editing.setStyle(value);
}
</script>

<template>
	<div
		class="pptx-vue-smartart-panel flex flex-col gap-2 text-xs"
		data-testid="smartart-panel"
		role="group"
		:aria-label="t('pptx.smartart.title')"
	>
		<p v-if="!isSmartArt" class="text-muted-foreground italic">
			{{ t('pptx.smartArt.selectPrompt') }}
		</p>

		<template v-else>
			<SmartArtLayoutSwitcher
				:current="editing.currentLayout.value"
				:can-edit="editable"
				@switch="onSwitchLayout"
			/>

			<label class="flex flex-col gap-1 text-[11px]">
				<span class="text-muted-foreground">{{ t('pptx.smartart.colorScheme') }}</span>
				<pptx-ui-select
					class="flex-1 bg-muted border border-border rounded px-1.5 py-0.5 w-full"
					data-testid="smartart-color-scheme"
					:aria-label="t('pptx.smartart.colorScheme')"
					:disabled="!editable"
					:value="editing.colorScheme.value"
					@change="onColorSchemeChange"
				>
					<option v-for="cs in colorSchemes" :key="cs" :value="cs">
						{{ colorSchemeLabel(cs) }}
					</option>
				</pptx-ui-select>
			</label>

			<div class="flex flex-col gap-1 text-[11px]">
				<span class="text-muted-foreground">{{ t('pptx.smartart.style') }}</span>
				<div class="flex gap-1" role="group" :aria-label="t('pptx.smartart.style')">
					<button
						v-for="s in styleOptions"
						:key="s"
						type="button"
						:disabled="!editable"
						:data-testid="`smartart-style-${s}`"
						:aria-pressed="editing.style.value === s"
						class="flex-1 px-2 py-1 text-[10px] rounded border transition-colors"
						:class="
							editing.style.value === s
								? 'border-primary bg-primary/20 text-primary'
								: 'border-border text-muted-foreground hover:bg-muted'
						"
						@click="onSetStyle(s)"
					>
						{{ styleLabel(s) }}
					</button>
				</div>
			</div>

			<pptx-ui-smartart-text-pane
				data-testid="smartart-text-pane"
				:state.prop="paneState"
				@pptx-smartart-data-change="onPaneChange"
			/>
		</template>
	</div>
</template>
