import { ELEMENT_ALT_TEXT_SECTIONS, getNonVisualDescriptionFields } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import React from 'react';

import { useTranslation } from '../../i18n/use-translation';
import { CARD, HEADING } from './inspector-pane-constants';
import { InspectorFields } from './InspectorFields';

interface AccessibilityTextSectionProps {
	selectedElement: PptxElement;
	canEdit: boolean;
	onUpdateElement: (updates: Partial<PptxElement>) => void;
}

/**
 * The Accessibility card of a shape, text box, connector or graphic frame:
 * the shared Alt Text descriptors (`ELEMENT_ALT_TEXT_SECTIONS`) drawn by the
 * generic renderer. A picture's alt text is drawn in `ImagePropertiesPanel`
 * from the same descriptors.
 */
export function AccessibilityTextSection({
	selectedElement,
	canEdit,
	onUpdateElement,
}: AccessibilityTextSectionProps): React.ReactElement | null {
	const { t } = useTranslation();
	const fields = getNonVisualDescriptionFields(selectedElement);
	if (!fields.showAltText && !fields.showTitle) {
		return null;
	}
	return (
		<div className={CARD} data-pptx-accessibility-text data-pptx-inspector-section='alt-text'>
			<div className={HEADING}>{t('pptx.accessibility.heading', 'Accessibility')}</div>
			<div className='grid grid-cols-2 gap-1.5 text-[11px]'>
				<InspectorFields
					sections={ELEMENT_ALT_TEXT_SECTIONS}
					target={selectedElement}
					disabled={!canEdit}
					onPatch={onUpdateElement}
				/>
			</div>
		</div>
	);
}
