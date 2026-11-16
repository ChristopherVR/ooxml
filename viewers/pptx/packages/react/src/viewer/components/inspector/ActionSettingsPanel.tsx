/**
 * ActionSettingsPanel: PowerPoint's Insert > Action dialog as an inspector card.
 *
 * An element carries two independent actions, one per trigger (`actionClick` /
 * `actionHover`). Each trigger is the shared `ACTION_SETTINGS_SECTIONS` drawn
 * by the generic renderer over the trigger's context (`actionSettingsTargetOf`:
 * its action, the pending pick and the deck's slides and custom shows); a
 * change holds its type as pending and, once it can carry its target, commits
 * through core's `elementActionToPptxAction`.
 */
import { ACTION_SETTINGS_SECTIONS, actionSettingsTargetOf } from 'ooxml-ui/pptx';
import type { ActionSettingsChange } from 'ooxml-ui/pptx';
import type { ElementActionType, PptxElement, PptxSlide } from 'pptx-viewer-core';
import { elementActionToPptxAction } from 'pptx-viewer-core';
import React, { useState } from 'react';

import { useTranslation } from '../../i18n/use-translation';
import { CARD, HEADING } from './inspector-pane-constants';
import { InspectorFields } from './InspectorFields';

interface ActionSettingsPanelProps {
	selectedElement: PptxElement;
	slides: PptxSlide[];
	canEdit: boolean;
	/** `data.customShows`, for the `customShow` target picker. */
	customShows: Array<{ id: string; name: string }>;
	onUpdateElement: (updates: Partial<PptxElement>) => void;
}

export function ActionSettingsPanel({
	selectedElement,
	slides,
	canEdit,
	customShows,
	onUpdateElement,
}: ActionSettingsPanelProps): React.ReactElement {
	const { t } = useTranslation();
	return (
		<div className={CARD} data-pptx-action-settings>
			<div className={HEADING}>{t('pptx.action.title', 'Action')}</div>
			<div className='space-y-2 text-[11px]'>
				{(['click', 'hover'] as const).map((trigger) => (
					<ActionTriggerSection
						// A pending pick belongs to the element it was made on: re-keying on
						// the element id drops it when the inspector moves on.
						key={`${trigger}-${selectedElement.id}`}
						trigger={trigger}
						label={
							trigger === 'click'
								? t('pptx.action.onClick', 'On Click')
								: t('pptx.action.onHover', 'On Hover')
						}
						element={selectedElement}
						slides={slides}
						customShows={customShows}
						canEdit={canEdit}
						onUpdateElement={onUpdateElement}
					/>
				))}
			</div>
		</div>
	);
}

interface ActionTriggerSectionProps {
	trigger: 'click' | 'hover';
	label: string;
	element: PptxElement;
	slides: PptxSlide[];
	customShows: Array<{ id: string; name: string }>;
	canEdit: boolean;
	onUpdateElement: (updates: Partial<PptxElement>) => void;
}

function ActionTriggerSection({
	trigger,
	label,
	element,
	slides,
	customShows,
	canEdit,
	onUpdateElement,
}: ActionTriggerSectionProps): React.ReactElement {
	const [pendingType, setPendingType] = useState<ElementActionType | undefined>(undefined);
	const target = actionSettingsTargetOf(element, trigger, pendingType, slides, customShows);
	const apply = (change: ActionSettingsChange): void => {
		setPendingType(change.pendingType);
		if (change.commit) {
			const action = elementActionToPptxAction(change.commit);
			onUpdateElement(
				(trigger === 'click'
					? { actionClick: action }
					: { actionHover: action }) as Partial<PptxElement>,
			);
		}
	};
	return (
		<div className='space-y-1.5' data-pptx-action-trigger={trigger}>
			<span className='text-muted-foreground font-medium'>{label}</span>
			<div className='grid grid-cols-2 gap-1.5'>
				<InspectorFields
					sections={ACTION_SETTINGS_SECTIONS}
					target={target}
					disabled={!canEdit}
					onPatch={apply}
				/>
			</div>
		</div>
	);
}
