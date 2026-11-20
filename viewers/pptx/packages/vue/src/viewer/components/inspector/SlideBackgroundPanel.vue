<script setup lang="ts">
import {
	inspectorFieldPatch,
	normalizeHexColor,
	resolveTemplateBackgroundRows,
	SLIDE_BACKGROUND_SECTIONS,
	slideBackgroundClearPatch,
	slideBackgroundClearState,
} from 'ooxml-ui/pptx';
/**
 * SlideBackgroundPanel: per-slide background editing (Vue port of React's
 * `SlideBackgroundPanel`). Edits the active slide's solid colour and
 * background image, and clears the background. In template-edit mode, also
 * shows the layout/master background card so their colour can be changed
 * without leaving the slide (the fuller Master Views overlay covers the same
 * ground but requires switching views).
 *
 * Emits `update` with a `Partial<PptxSlide>` patch; the host (SlideInspector ->
 * PowerPointViewer) applies it to the active slide with history.
 */
import type { PptxSlide, PptxSlideMaster } from 'pptx-viewer-core';
import { computed, ref } from 'vue';
import { useI18n } from 'vue-i18n';

import { useLinkedContentUrl } from '../../composables/linked-content';
import { injectRecentColors } from '../../composables/recent-colors-context';
import RecentColorsRow from '../RecentColorsRow.vue';
import { CARD, HEADING } from './inspector-cards';
import InspectorFields from './InspectorFields.vue';

const props = withDefaults(
	defineProps<{
		slide: PptxSlide | undefined;
		canEdit?: boolean;
		editTemplateMode?: boolean;
		slideMasters?: PptxSlideMaster[];
		getTemplateBackgroundColor?: (path: string) => string | undefined;
	}>(),
	{ canEdit: true },
);

const emit = defineEmits<{
	update: [patch: Partial<PptxSlide>];
	'set-template-background': [path: string, backgroundColor: string];
}>();

const { t } = useI18n();
const recentColors = injectRecentColors();

const templateRows = computed(() =>
	props.slide
		? resolveTemplateBackgroundRows(
				props.slide,
				props.slideMasters,
				t('pptx.master.layout'),
				t('pptx.master.master'),
			)
		: {},
);

function templateBackgroundValue(path: string): string {
	return normalizeHexColor(props.getTemplateBackgroundColor?.(path), '#ffffff');
}

function onTemplateColorChange(path: string, event: Event): void {
	const hex = (event.target as HTMLInputElement).value;
	emit('set-template-background', path, hex);
	recentColors?.push(hex);
}

const fileInput = ref<HTMLInputElement | null>(null);

const backgroundImage = computed(() => props.slide?.backgroundImage);
/** The preview draws a linked background only once this viewer allows linked content. */
const backgroundPreview = useLinkedContentUrl(() => backgroundImage.value);
const clearState = computed(() => slideBackgroundClearState(props.slide, props.canEdit));

/** A "Recent colours" pick: the same patch as the shared colour field. */
function onColorPick(hex: string): void {
	const patch = inspectorFieldPatch(
		SLIDE_BACKGROUND_SECTIONS,
		'slideBackground',
		'backgroundColor',
		props.slide,
		hex,
	);
	if (patch) {
		emit('update', patch);
	}
	recentColors?.push(hex);
}

function onImageChange(event: Event): void {
	const input = event.target as HTMLInputElement;
	const file = input.files?.[0];
	input.value = '';
	if (!file) {
		return;
	}
	const reader = new FileReader();
	reader.onload = () => {
		if (typeof reader.result === 'string') {
			emit('update', { backgroundImage: reader.result });
		}
	};
	reader.readAsDataURL(file);
}

function removeImage(): void {
	emit('update', { backgroundImage: undefined });
}

function clearBackground(): void {
	emit('update', slideBackgroundClearPatch());
}
</script>

<template>
	<div :class="CARD">
		<div :class="HEADING">{{ t('pptx.viewer.background') }}</div>
		<InspectorFields
			:sections="SLIDE_BACKGROUND_SECTIONS"
			:target="slide"
			:disabled="!canEdit"
			@patch="emit('update', $event)"
		/>
		<RecentColorsRow v-if="recentColors" :colors="recentColors.recent.value" @pick="onColorPick" />

		<div class="space-y-1">
			<div class="flex items-center gap-2 text-[11px]">
				<span class="w-10 shrink-0 text-muted-foreground">{{
					t('pptx.slideBackground.image')
				}}</span>
				<input
					ref="fileInput"
					type="file"
					accept="image/png,image/jpeg,image/gif,image/webp,image/svg+xml"
					class="hidden"
					:disabled="!canEdit"
					@change="onImageChange"
				/>
				<button
					type="button"
					class="flex-1 rounded border border-border bg-muted px-2 py-1 text-center text-[11px] hover:bg-accent disabled:opacity-50"
					:disabled="!canEdit"
					@click="fileInput?.click()"
				>
					{{
						backgroundImage
							? t('pptx.slideBackground.replaceImage')
							: t('pptx.slideBackground.chooseImage')
					}}
				</button>
			</div>
			<div v-if="backgroundImage" class="relative mt-1">
				<img
					v-if="backgroundPreview"
					:src="backgroundPreview"
					:alt="t('pptx.slideBackground.backgroundPreview')"
					class="h-16 w-full rounded border border-border object-cover"
				/>
				<button
					type="button"
					class="absolute right-0.5 top-0.5 rounded bg-background/80 p-0.5 text-[10px] transition-colors hover:bg-red-700 disabled:opacity-50"
					:disabled="!canEdit"
					:title="t('pptx.slideBackground.removeBackgroundImage')"
					:aria-label="t('pptx.slideBackground.removeBackgroundImage')"
					@click="removeImage"
				>
					X
				</button>
			</div>
		</div>

		<button
			v-if="clearState.visible"
			type="button"
			class="w-full rounded border border-border bg-muted px-2 py-1 text-center text-[11px] text-red-400 hover:text-red-300 disabled:opacity-50"
			:disabled="!clearState.enabled"
			@click="clearBackground"
		>
			{{ t('pptx.slideBackground.clearBackground') }}
		</button>
	</div>

	<div v-if="editTemplateMode && (templateRows.layout || templateRows.master)" :class="CARD">
		<div :class="HEADING">{{ t('pptx.slideBackground.templateBackgroundsHeading') }}</div>

		<label v-if="templateRows.layout" class="flex items-center gap-2 text-[11px]">
			<span class="w-14 shrink-0 truncate text-muted-foreground" :title="templateRows.layout.title">
				{{ t('pptx.master.layout') }}
			</span>
			<input
				type="color"
				:value="templateBackgroundValue(templateRows.layout.path)"
				:disabled="!canEdit"
				class="h-6 w-8 cursor-pointer rounded border border-border bg-muted"
				@change="onTemplateColorChange(templateRows.layout.path, $event)"
			/>
			<span class="truncate text-[10px] text-muted-foreground">{{
				templateRows.layout.label
			}}</span>
		</label>

		<label v-if="templateRows.master" class="flex items-center gap-2 text-[11px]">
			<span class="w-14 shrink-0 truncate text-muted-foreground" :title="templateRows.master.title">
				{{ t('pptx.master.master') }}
			</span>
			<input
				type="color"
				:value="templateBackgroundValue(templateRows.master.path)"
				:disabled="!canEdit"
				class="h-6 w-8 cursor-pointer rounded border border-border bg-muted"
				@change="onTemplateColorChange(templateRows.master.path, $event)"
			/>
			<span class="truncate text-[10px] text-muted-foreground">{{
				templateRows.master.label
			}}</span>
		</label>
	</div>
</template>
