<script setup lang="ts">
import { SLIDE_SHOW_COMMAND_GROUPS, SLIDE_SHOW_OPTIONS } from 'ooxml-ui/pptx';
import type {
	RibbonCommandRequestEvent,
	RibbonControlId,
	SlideShowOptionsChangeEvent,
} from 'ooxml-ui/pptx';
import { useI18n } from 'vue-i18n';

import { useToolbarVisibility } from '../../composables/useToolbarVisibility';
import type { SlideShowSectionProps } from './slide-show-section-props';
import SubtitleSettingsControl from './SubtitleSettingsControl.vue';

const props = defineProps<SlideShowSectionProps>();
const { t } = useI18n();
const { isHidden } = useToolbarVisibility(() => props.hiddenActions);
const optionLabels = () =>
	Object.fromEntries(SLIDE_SHOW_OPTIONS.map((option) => [option.id, t(option.labelKey)]));
function commitOptions(event: Event): void {
	props.onPresentationPropertiesChange?.((event as SlideShowOptionsChangeEvent).detail);
}
function requestCommand(event: Event): void {
	const actions: Partial<Record<RibbonControlId, () => void>> = {
		'slideShow.startSlideShow.fromBeginning': props.onPresentFromBeginning,
		'slideShow.startSlideShow.fromCurrent': props.onPresent,
		'slideShow.present.presenterView': props.onEnterPresenterView,
		'slideShow.startSlideShow.customShow': () => props.customShowControls.onCreateCustomShow(),
		'slideShow.present.broadcast': props.onOpenBroadcastDialog,
		'slideShow.setUp.setUpSlideShow': props.onOpenSetUpSlideShow,
		'slideShow.setUp.hideSlide': props.onToggleHideSlide,
		'slideShow.setUp.rehearseTimings': props.onEnterRehearsalMode,
		'slideShow.setUp.record': props.onEnterRehearsalMode,
	};
	actions[(event as RibbonCommandRequestEvent).detail.id]?.();
}
</script>

<template>
	<pptx-ui-ribbon-group
		v-for="group in SLIDE_SHOW_COMMAND_GROUPS"
		:key="group.id"
		:label="t(group.labelKey)"
		:data-ribbon-group="group.id"
	>
		<template v-for="command in group.commands" :key="command.id">
			<pptx-ui-ribbon-command
				v-if="command.id !== 'slideShow.present.broadcast' || !isHidden('broadcast')"
				:data-ribbon-control="command.id"
				:label="t(command.labelKey)"
				:icon="command.icon"
				:title="t(command.tooltipKey ?? command.labelKey)"
				:disabled="command.unsupported ? '' : undefined"
				:active="
					command.id === 'slideShow.setUp.hideSlide' && props.activeSlideHidden ? '' : undefined
				"
				:pressed="
					command.id === 'slideShow.setUp.hideSlide' ? String(props.activeSlideHidden) : undefined
				"
				@command-request="requestCommand"
			/>
		</template>
	</pptx-ui-ribbon-group>
	<pptx-ui-ribbon-group :label="t('pptx.slideShow.options')">
		<pptx-ui-slide-show-options
			:presentationProperties.prop="props.presentationProperties"
			:labels.prop="optionLabels()"
			@show-options-change="commitOptions"
		>
			<div class="contents" data-ribbon-group="slideShow.captions">
				<pptx-ui-ribbon-toggle
					data-ribbon-control="slideShow.captions.subtitles"
					:label="t('pptx.slideShow.subtitles')"
					:title="t('pptx.slideShow.subtitlesTooltip')"
					:checked="props.showSubtitles ? '' : undefined"
					@toggle-request="props.onToggleSubtitles()"
				/>
				<SubtitleSettingsControl />
			</div>
		</pptx-ui-slide-show-options>
	</pptx-ui-ribbon-group>
</template>
