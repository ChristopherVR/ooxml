<script setup lang="ts">
import { describeCollaborationShellState } from 'pptx-vue-viewer';
import {
	CollaborationCursors,
	InlineTextEditor,
	overlayInlineTextSnapshot,
	RemoteSelectionOverlay,
	SelectionOverlay,
	SlideCanvas,
	SlidesPane,
	useCollaboration,
	useEditorHistory,
	useEditorOperations,
	useElementDrag,
	useInlineEditing,
	useLoadContent,
	useSectionOperations,
	useSlideOperations,
} from 'pptx-vue-viewer/viewer';
import type { UseCollaborationResult, UseInlineEditingResult } from 'pptx-vue-viewer/viewer';
import { computed, onBeforeUnmount, ref, watch } from 'vue';
import { useI18n } from 'vue-i18n';

import type { HostOwnedDemo } from '../../shared/host-owned-collaboration';

// The slide rail's chrome classes and tokens come from the full stylesheet.
import 'pptx-vue-viewer/styles';

const props = defineProps<{ host: HostOwnedDemo }>();
const root = ref<HTMLElement>();
const canvas = ref<InstanceType<typeof SlideCanvas>>();
const scale = ref(1);
const current = ref(0);
const loadVersion = ref(0);
const content = useLoadContent(() => props.host.source, {
	getPendingInlineEdit: () => {
		if (!shellState.value.canEdit) {
			return undefined;
		}
		if (edit.isInlineInputPending()) {
			throw new Error('Finish the current text input before saving.');
		}
		const snapshot = edit.readInlineSnapshot();
		return snapshot && ops.activeSlide.value
			? { snapshot, target: { slideId: ops.activeSlide.value.id } }
			: undefined;
	},
	onContentApplied: () => {
		loadVersion.value++;
	},
});
const { slides, canvasSize, mediaDataUrls, templateElementsBySlideId } = content;
const history = useEditorHistory(slides);
const ops = useEditorOperations({
	slides,
	activeSlideIndex: current,
	pushHistory: history.pushHistory,
	templateElementsBySlideId,
});
// The rail's commands: plain edits of `slides`, which the collaboration
// session below publishes to peers like any other edit.
const slideOps = useSlideOperations({
	slides,
	activeSlideIndex: current,
	pushHistory: history.pushHistory,
	templateElementsBySlideId,
});
const sectionOps = useSectionOperations({
	sections: content.sections,
	slides,
	activeSlideIndex: current,
	pushHistory: history.pushHistory,
});
const collaboration: UseCollaborationResult = useCollaboration({
	slides,
	collaboration: () => props.host.config,
	canEdit: () => props.host.editable,
	sourcePending: content.loading,
	sourceError: () => Boolean(content.error.value || content.isEncrypted.value),
	onRemoteSlides: (next) => {
		slides.value = next;
	},
	loadVersion,
	getLoadOrigin: () => 'bootstrap',
	canvasWidth: computed(() => canvasSize.value.width),
	canvasHeight: computed(() => canvasSize.value.height),
	serialize: () => content.getContent(),
});
const { shellState, remotePresences, cursors } = collaboration;
const edit: UseInlineEditingResult = useInlineEditing({
	canEdit: () => shellState.value.canEdit,
	findActiveElement: (id) => ops.activeSlide.value?.elements.find((element) => element.id === id),
	ops,
	livePatcher: () => collaboration.livePatcher,
	activeSlide: () => ops.activeSlide.value,
	onConnectedSuspend: (snapshot) => {
		const id = ops.activeSlide.value?.id;
		const current = slides.value;
		const next = current.map((slide) => {
			if (slide.id !== id) return slide;
			const elements = overlayInlineTextSnapshot(slide.elements, snapshot);
			return elements === slide.elements ? slide : { ...slide, elements: [...elements] };
		});
		if (next.some((slide, index) => slide !== current[index])) slides.value = next;
	},
});
const { inlineEditingElement, inlineEditingElementId } = edit;
const drag = useElementDrag({
	findActiveElement: (id) => ops.activeSlide.value?.elements.find((element) => element.id === id),
	pushHistory: history.pushHistory,
	effectiveZoom: computed(() => scale.value),
	activeTemplateElements: computed(
		() => templateElementsBySlideId.value[ops.activeSlide.value?.id ?? ''] ?? [],
	),
	activeSlide: ops.activeSlide,
	activeSlideIndex: current,
	slides,
	templateElementsBySlideId,
	canvasSize,
	enterInlineEdit: enterText,
});
onBeforeUnmount(() => drag.cancelElementDrag());
watch(
	() => shellState.value.canEdit,
	(allowed) => {
		// Connected retirement is owned by useInlineEditing, without a history commit.
		if (!allowed) {
			drag.cancelElementDrag();
			edit.commitInlineEdit();
		}
	},
	{ flush: 'sync' },
);
watch(current, (index) => {
	edit.commitInlineEdit();
	collaboration.setActiveSlide(index);
});
watch(ops.selectedElementIds, (ids) => collaboration.setSelection(ids));
const activeSlide = ops.activeSlide;
const { t } = useI18n();
const statusText = computed(() => describeCollaborationShellState(shellState.value, t));

function pick(event: MouseEvent): string | undefined {
	return (
		(event.target as HTMLElement).closest('[data-element-id]')?.getAttribute('data-element-id') ??
		undefined
	);
}
function select(event: PointerEvent): void {
	if (!shellState.value.canEdit || (event.target as HTMLElement).closest('[data-inline-editor]')) {
		return;
	}
	edit.commitInlineEdit();
	if (edit.isInlineInputPending()) return;
	const id = pick(event);
	ops.selectedElementIds.value = id ? [id] : [];
	root.value?.focus();
	if (id) {
		event.preventDefault();
		drag.startElementDrag(id, event, false);
	}
}
function openText(event: MouseEvent): void {
	const id = pick(event);
	if (id) enterText(id);
}
function enterText(id: string): void {
	if (!shellState.value.canEdit) return;
	edit.commitInlineEdit();
	if (!edit.isInlineInputPending()) edit.enterInlineEdit(id);
}
function navigate(index: number): void {
	edit.commitInlineEdit();
	if (!edit.isInlineInputPending()) current.value = index;
}
/** Runs a rail command once any open text edit is committed, and only while editing is allowed. */
function railEdit<A extends unknown[]>(run: (...args: A) => void): (...args: A) => void {
	return (...args) => {
		if (!shellState.value.canEdit) return;
		edit.commitInlineEdit();
		if (!edit.isInlineInputPending()) run(...args);
	};
}
const rail = {
	moveSlide: railEdit(slideOps.moveSlide),
	addSlide: railEdit(() => slideOps.addSlide()),
	addSlideAfter: railEdit((index: number) => slideOps.addSlide(index)),
	duplicateSlide: railEdit(slideOps.duplicateSlide),
	deleteSlide: railEdit(slideOps.deleteSlide),
	toggleSlideHidden: railEdit((index: number) => {
		const target = slides.value[index];
		if (!target) return;
		history.pushHistory();
		slides.value = slides.value.map((slide, i) =>
			i === index ? { ...slide, hidden: !slide.hidden } : slide,
		);
	}),
	addSection: railEdit(sectionOps.addSection),
	renameSection: railEdit(sectionOps.renameSection),
	deleteSection: railEdit(sectionOps.deleteSection),
	moveSectionUp: railEdit(sectionOps.moveSectionUp),
	moveSectionDown: railEdit(sectionOps.moveSectionDown),
};
function moveCursor(event: PointerEvent): void {
	const stage = canvas.value?.getStageElement();
	if (!stage) {
		return;
	}
	const rect = stage.getBoundingClientRect();
	collaboration.setCursor(
		(event.clientX - rect.left) / scale.value,
		(event.clientY - rect.top) / scale.value,
	);
}
function onKeydown(event: KeyboardEvent): void {
	if (!shellState.value.canEdit || inlineEditingElementId.value) {
		return;
	}
	const element = activeSlide.value?.elements.find(
		(item) => item.id === ops.selectedElementIds.value[0],
	);
	if (!element || !['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown'].includes(event.key)) {
		return;
	}
	event.preventDefault();
	ops.moveElement(element.id, {
		x: element.x + (event.key === 'ArrowRight' ? 1 : event.key === 'ArrowLeft' ? -1 : 0),
		y: element.y + (event.key === 'ArrowDown' ? 1 : event.key === 'ArrowUp' ? -1 : 0),
	});
}
defineExpose({
	getContent: async () => {
		if (!collaboration.livePatcher.isActive()) edit.commitInlineEdit();
		return content.getContent();
	},
	setScale: (next: number) => {
		scale.value = next;
	},
});
</script>

<template>
	<div ref="root" data-host-custom-shell="vue" class="shell" tabindex="0" @keydown="onKeydown">
		<nav class="custom-nav" aria-label="Custom slide navigation">
			<button v-for="(_, index) in slides" :key="index" @click="navigate(index)">
				Slide {{ index + 1 }}
			</button>
			<output :aria-label="t('pptx.collaboration.shellStatusLabel')">{{ statusText }}</output>
		</nav>
		<p v-if="content.error.value" role="alert">{{ content.error.value }}</p>
		<div class="workspace">
			<SlidesPane
				:slides="slides"
				:template-elements-by-slide-id="templateElementsBySlideId"
				:active-slide-index="current"
				:canvas-size="canvasSize"
				:media-data-urls="mediaDataUrls"
				:section-groups="sectionOps.slidesBySection.value"
				:can-edit="shellState.canEdit"
				:header-footer="content.headerFooter.value"
				:custom-properties="content.customProperties.value"
				:first-slide-number="content.firstSlideNumber.value"
				:theme="content.theme.value"
				:table-style-map="content.tableStyleMap.value"
				@select="navigate"
				@move-slide="rail.moveSlide"
				@add-slide="rail.addSlide"
				@add-slide-after="rail.addSlideAfter"
				@duplicate-slide="rail.duplicateSlide"
				@delete-slide="rail.deleteSlide"
				@toggle-slide-hidden="rail.toggleSlideHidden"
				@add-section="rail.addSection"
				@rename-section="rail.renameSection"
				@delete-section="rail.deleteSection"
				@move-section-up="rail.moveSectionUp"
				@move-section-down="rail.moveSectionDown"
				@toggle-section-collapse="sectionOps.toggleSectionCollapse"
			/>
			<div class="stage">
				<SlideCanvas
					ref="canvas"
					:slide="activeSlide"
					:canvas-size="canvasSize"
					:media-data-urls="mediaDataUrls"
					:zoom="scale"
					:template-elements="activeSlide ? templateElementsBySlideId[activeSlide.id] : []"
					:inline-editing-element-id="inlineEditingElementId"
					@pointerdown="select"
					@pointermove="moveCursor"
					@dblclick="openText"
				>
					<SelectionOverlay
						v-if="shellState.canEdit && !inlineEditingElement"
						:elements="activeSlide?.elements ?? []"
						:selected-ids="ops.selectedElementIds.value"
						:zoom="scale"
						@transform-start="drag.onTransformStart"
						@transform="drag.onTransform"
						@transform-end="drag.onTransformEnd"
						@adjust-start="drag.onAdjustStart"
						@adjust="drag.onAdjust"
						@adjust-end="drag.onAdjustEnd"
						@request-edit="({ id }) => enterText(id)"
					/>
					<InlineTextEditor
						v-if="shellState.canEdit && inlineEditingElement"
						:key="inlineEditingElement.id"
						:element="inlineEditingElement"
						:live-patcher="collaboration.livePatcher"
						:slide-id="activeSlide?.id"
						@change="edit.updateInlineText"
						@commit="edit.commitInlineEdit"
						@cancel="edit.cancelInlineEdit"
						@list-session="edit.onListSession"
					/>
					<RemoteSelectionOverlay
						:presences="remotePresences"
						:elements="activeSlide?.elements ?? []"
						:active-slide-index="current"
					/>
					<CollaborationCursors :cursors="cursors" />
				</SlideCanvas>
			</div>
		</div>
	</div>
</template>

<style scoped>
.shell {
	height: 100%;
	display: flex;
	flex-direction: column;
	background: #e2e8f0;
	color: #0f172a;
}
/* A class, not `nav`: scoped styles also reach the root of the rail. */
.custom-nav {
	display: flex;
	gap: 8px;
	padding: 8px;
}
.workspace {
	flex: 1;
	min-height: 0;
	display: flex;
}
.stage {
	flex: 1;
	min-width: 0;
	min-height: 0;
	position: relative;
	display: flex;
	flex-direction: column;
}
</style>
