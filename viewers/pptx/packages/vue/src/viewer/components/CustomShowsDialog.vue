<script setup lang="ts">
/**
 * CustomShowsDialog: Slide Show > Custom Slide Show > Custom Shows.
 *
 * New, Copy and Remove for the shows and, for the selected one, PowerPoint's two lists
 * (slides in the presentation, slides in the custom show) with Add, Remove and Move Up/Down.
 * Every decision is the shared `createCustomShowsController`; this file only draws it.
 *
 * Emits `save` (the shows, plus slides with jumps into removed shows cleared, or undefined),
 * `select` (the show playback is restricted to; null for All Slides), `run` (the Show button,
 * after the dialog committed) and `close`.
 */
import { createCustomShowsController } from 'ooxml-ui/pptx';
import type { PptxCustomShow, PptxSlide } from 'pptx-viewer-core';
import { computed, ref } from 'vue';
import { useI18n } from 'vue-i18n';

import { useController } from '../composables/useController';
import DialogFooter from './DialogFooter.vue';
import ModalDialog from './ModalDialog.vue';

const props = defineProps<{
	open: boolean;
	customShows: PptxCustomShow[];
	slides: PptxSlide[];
	/** The show playback is restricted to, or null for the whole deck. */
	activeShowId: string | null;
}>();

const emit = defineEmits<{
	save: [shows: PptxCustomShow[], clearedSlides: PptxSlide[] | undefined];
	select: [showId: string | null];
	run: [showId: string];
	close: [];
}>();

const { t } = useI18n();
// A fresh draft per mount: the parent renders this with `v-if`, so each open starts clean.
const controller = createCustomShowsController(props.customShows, props.activeShowId);
const snapshot = useController(controller);
const view = computed(() => {
	void snapshot.value;
	return controller.view(props.slides);
});
const active = ref<string>(props.activeShowId ?? '');
const pickerValue = computed(() =>
	view.value.shows.some(({ id }) => id === active.value) ? active.value : '',
);

function setActive(next: string): void {
	active.value = next;
	emit('select', next || null);
}

function commit(): PptxCustomShow[] {
	const result = controller.result(props.slides);
	emit('save', result.shows, result.slides);
	emit('select', result.shows.some(({ id }) => id === active.value) ? active.value : null);
	return result.shows;
}

function onFooter(id: string): void {
	if (id === 'save') {
		commit();
		emit('close');
	} else if (id === 'show') {
		const selectedId = controller.getSnapshot().selectedId;
		const show = commit().find((candidate) => candidate.id === selectedId);
		if (show) {
			emit('run', show.id);
			emit('close');
		}
	} else {
		emit('close');
	}
}

const footerActions = computed(() => [
	{ id: 'show', label: t('pptx.customShows.showButton') },
	{ id: 'cancel', label: t('pptx.customShows.cancel') },
	{ id: 'save', label: t('pptx.common.ok'), variant: 'primary' as const },
]);
</script>

<template>
	<ModalDialog :open="props.open" :title="t('pptx.customShows.title')" @close="emit('close')">
		<div class="pptx-vue-cs-dialog flex flex-col gap-3 text-xs text-foreground">
			<label class="pptx-vue-cs-active flex items-center gap-2">
				<span>{{ t('pptx.customShows.selectCustomShow') }}</span>
				<pptx-ui-select
					class="pptx-vue-cs-select min-w-0 flex-1"
					:aria-label="t('pptx.customShows.selectCustomShow')"
					:value="pickerValue"
					@change="setActive(($event.target as HTMLSelectElement).value)"
				>
					<option value="">{{ t('pptx.customShows.allSlides') }}</option>
					<option v-for="show in view.shows" :key="show.id" :value="show.id">
						{{ show.name }}
					</option>
				</pptx-ui-select>
			</label>

			<div class="pptx-vue-cs-body">
				<div class="pptx-vue-cs-shows flex flex-col gap-1.5">
					<ul
						role="listbox"
						:aria-label="t('pptx.customShows.savedShows')"
						class="pptx-vue-cs-list"
					>
						<li
							v-for="show in view.shows"
							:key="show.id"
							role="option"
							tabindex="0"
							:aria-selected="show.selected"
							@click="controller.select(show.id)"
							@keydown.enter="controller.select(show.id)"
						>
							{{ show.name }}
						</li>
						<li v-if="view.shows.length === 0" class="pptx-vue-cs-empty">
							{{ t('pptx.customShows.empty') }}
						</li>
					</ul>
					<button
						type="button"
						class="pptx-vue-cs-btn"
						@click="controller.create(t('pptx.customShows.newDefaultName'))"
					>
						{{ t('pptx.customShows.createNew') }}
					</button>
					<button
						type="button"
						class="pptx-vue-cs-btn"
						:disabled="!view.canEdit"
						@click="controller.copySelected()"
					>
						{{ t('pptx.customShows.copy') }}
					</button>
					<button
						type="button"
						class="pptx-vue-cs-btn"
						:disabled="!view.canEdit"
						@click="controller.removeSelected()"
					>
						{{ t('pptx.customShows.remove') }}
					</button>
				</div>

				<div v-if="view.selected" class="pptx-vue-cs-editor flex flex-col gap-2">
					<label class="flex items-center gap-2">
						<span>{{ t('pptx.customShows.renameNameLabel') }}</span>
						<input
							class="pptx-vue-cs-input min-w-0 flex-1"
							type="text"
							:aria-label="t('pptx.customShows.renameNameLabel')"
							:value="view.selected.name"
							@change="controller.rename(($event.target as HTMLInputElement).value)"
						/>
					</label>
					<div class="pptx-vue-cs-lists">
						<div>
							<strong>{{ t('pptx.customShows.slidesInPresentation') }}</strong>
							<ul
								role="listbox"
								aria-multiselectable="true"
								:aria-label="t('pptx.customShows.slidesInPresentation')"
								class="pptx-vue-cs-list"
							>
								<li
									v-for="row in view.presentationSlides"
									:key="row.rId"
									role="option"
									tabindex="0"
									:aria-selected="row.selected"
									@click="controller.toggleLeft(row.rId)"
									@dblclick="controller.addSlide(row.rId)"
									@keydown.space.prevent="controller.toggleLeft(row.rId)"
								>
									{{ row.number }}. {{ row.title }}
								</li>
							</ul>
						</div>
						<div class="pptx-vue-cs-actions">
							<button
								type="button"
								class="pptx-vue-cs-btn"
								:disabled="!view.canAdd"
								@click="controller.addSelected(props.slides)"
							>
								{{ t('pptx.customShows.add') }}
							</button>
							<button
								type="button"
								class="pptx-vue-cs-btn"
								:disabled="!view.canRemoveSlides"
								@click="controller.removeSelectedSlides()"
							>
								{{ t('pptx.customShows.remove') }}
							</button>
						</div>
						<div>
							<strong>{{ t('pptx.customShows.slidesInShow') }}</strong>
							<ul
								role="listbox"
								aria-multiselectable="true"
								:aria-label="t('pptx.customShows.slidesInShow')"
								class="pptx-vue-cs-list"
							>
								<li
									v-for="row in view.showSlides"
									:key="row.position"
									role="option"
									tabindex="0"
									:aria-selected="row.selected"
									@click="controller.toggleRight(row.position ?? 0)"
									@keydown.space.prevent="controller.toggleRight(row.position ?? 0)"
								>
									{{ row.number }}. {{ row.title }}
								</li>
								<li v-if="view.showSlides.length === 0" class="pptx-vue-cs-empty">
									{{ t('pptx.customShows.noSlidesYet') }}
								</li>
							</ul>
						</div>
						<div class="pptx-vue-cs-actions">
							<button
								type="button"
								class="pptx-vue-cs-btn"
								:disabled="!view.canMoveUp"
								@click="controller.moveUp()"
							>
								{{ t('pptx.customShows.moveUp') }}
							</button>
							<button
								type="button"
								class="pptx-vue-cs-btn"
								:disabled="!view.canMoveDown"
								@click="controller.moveDown()"
							>
								{{ t('pptx.customShows.moveDown') }}
							</button>
						</div>
					</div>
				</div>
			</div>
		</div>
		<template #footer>
			<DialogFooter :actions="footerActions" @action="onFooter" />
		</template>
	</ModalDialog>
</template>

<style scoped>
.pptx-vue-cs-body {
	display: grid;
	grid-template-columns: 190px 1fr;
	gap: 12px;
	min-width: min(760px, 86vw);
}
.pptx-vue-cs-lists {
	display: grid;
	grid-template-columns: 1fr auto 1fr auto;
	gap: 8px;
}
.pptx-vue-cs-actions {
	display: flex;
	flex-direction: column;
	justify-content: center;
	gap: 6px;
}
.pptx-vue-cs-list {
	margin: 0;
	padding: 0;
	max-height: 230px;
	overflow: auto;
	list-style: none;
	border: 1px solid var(--pptx-border, #d1d5db);
	border-radius: 4px;
}
.pptx-vue-cs-list li {
	padding: 5px 8px;
	cursor: pointer;
}
.pptx-vue-cs-list li[aria-selected='true'] {
	background: var(--pptx-primary, #2563eb);
	color: #fff;
}
.pptx-vue-cs-list li.pptx-vue-cs-empty {
	cursor: default;
	opacity: 0.7;
}
.pptx-vue-cs-btn,
.pptx-vue-cs-input {
	border: 1px solid var(--pptx-border, #d1d5db);
	border-radius: 4px;
	padding: 5px 8px;
	background: var(--pptx-muted, #f3f4f6);
	color: inherit;
}
.pptx-vue-cs-btn:disabled {
	opacity: 0.5;
}
@media (max-width: 700px) {
	.pptx-vue-cs-body,
	.pptx-vue-cs-lists {
		grid-template-columns: 1fr;
		min-width: 0;
	}
}
</style>
