import type { ToolbarActionId, RibbonControlId } from 'ooxml-ui/pptx';
import { SLIDE_SHOW_COMMAND_GROUPS } from 'ooxml-ui/pptx';
import type { PptxPresentationProperties } from 'pptx-viewer-core';
import React from 'react';

import { useToolbarVisibility } from '../../hooks/useToolbarVisibility';
import { useTranslation } from '../../i18n/use-translation';
import type { ViewerMode } from '../../types';
import { RibbonGroupScope } from './PowerPointRibbonControls';
import { SlideShowOptions } from './SlideShowOptions';
import { SubtitleSettingsControl } from './SubtitleSettingsControl';
import { WebRibbonCommand, WebRibbonGroup, WebRibbonToggle } from './WebRibbonControls';

export interface SlideShowSectionProps {
	onPresent: () => void;
	/**
	 * "From Beginning": enters the show on its first slide regardless of the
	 * active slide. Falls back to `onSetMode('present')` when omitted.
	 */
	onPresentFromBeginning?: () => void;
	onEnterPresenterView: () => void;
	onEnterRehearsalMode: () => void;
	onOpenSetUpSlideShow: () => void;
	/** Slide Show > Custom Slide Show: open the Custom Shows dialog. */
	onOpenCustomShows: () => void;
	/**
	 * PowerPoint's Hide Slide toggle: marks the ACTIVE slide to be skipped during
	 * the show while leaving it in the deck, the thumbnail rail and the sorter.
	 */
	onToggleHideSlide: () => void;
	/** Whether the active slide is currently hidden, for the toggle's pressed state. */
	activeSlideHidden: boolean;
	onOpenBroadcastDialog: () => void;
	onToggleSubtitles: () => void;
	showSubtitles: boolean;
	onSetMode: (mode: ViewerMode) => void;
	/** Host-supplied list of toolbar buttons/ribbon tabs to hide. */
	hiddenActions?: readonly ToolbarActionId[];
	/** Deck presentation properties backing the Options checkboxes. */
	presentationProperties?: PptxPresentationProperties;
	/** Commit an Options checkbox onto the deck's presentation properties. */
	onPresentationPropertiesChange?: (updates: Partial<PptxPresentationProperties>) => void;
}

export function SlideShowSection(p: SlideShowSectionProps): React.ReactElement {
	const { t } = useTranslation();
	const { isHidden } = useToolbarVisibility(p.hiddenActions);
	const actions: Partial<Record<RibbonControlId, () => void>> = {
		'slideShow.startSlideShow.fromBeginning':
			p.onPresentFromBeginning ?? (() => p.onSetMode('present')),
		'slideShow.startSlideShow.fromCurrent': p.onPresent,
		'slideShow.present.presenterView': p.onEnterPresenterView,
		'slideShow.startSlideShow.customShow': p.onOpenCustomShows,
		'slideShow.present.broadcast': p.onOpenBroadcastDialog,
		'slideShow.setUp.setUpSlideShow': p.onOpenSetUpSlideShow,
		'slideShow.setUp.hideSlide': p.onToggleHideSlide,
		'slideShow.setUp.rehearseTimings': p.onEnterRehearsalMode,
		'slideShow.setUp.record': p.onEnterRehearsalMode,
	};
	const request = (id: RibbonControlId) => actions[id]?.();
	return (
		<>
			{SLIDE_SHOW_COMMAND_GROUPS.map((group) => (
				<WebRibbonGroup key={group.id} groupId={group.id} label={t(group.labelKey)}>
					{group.commands
						.filter(
							(command) => command.id !== 'slideShow.present.broadcast' || !isHidden('broadcast'),
						)
						.map((command) => {
							const hideSlide = command.id === 'slideShow.setUp.hideSlide';
							const button = (
								<WebRibbonCommand
									key={command.id}
									controlId={command.id}
									label={t(command.labelKey)}
									icon={command.icon}
									title={t(command.tooltipKey ?? command.labelKey)}
									disabled={command.unsupported}
									active={hideSlide ? p.activeSlideHidden : false}
									pressed={hideSlide ? p.activeSlideHidden : undefined}
									onCommand={request}
								/>
							);
							return button;
						})}
				</WebRibbonGroup>
			))}
			<WebRibbonGroup label={t('pptx.slideShow.options')}>
				<SlideShowOptions
					presentationProperties={p.presentationProperties}
					onChange={p.onPresentationPropertiesChange}
				>
					<RibbonGroupScope id='slideShow.captions'>
						<WebRibbonToggle
							controlId='slideShow.captions.subtitles'
							label={t('pptx.slideShow.subtitles')}
							checked={p.showSubtitles}
							title={t('pptx.slideShow.subtitlesTooltip')}
							onToggle={() => p.onToggleSubtitles()}
						/>
						<SubtitleSettingsControl />
					</RibbonGroupScope>
				</SlideShowOptions>
			</WebRibbonGroup>
		</>
	);
}
