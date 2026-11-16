import {
	ELEMENT_ALT_TEXT_SECTIONS,
	imageResetPatch,
	imageResetState,
	PICTURE_ADJUSTMENT_SECTIONS,
	PICTURE_ARTISTIC_SECTIONS,
	PICTURE_EFFECT_SECTIONS,
} from 'ooxml-ui/pptx';
import type { PptxElement, PptxImageLikeElement } from 'pptx-viewer-core';
import React from 'react';

import { useTranslation } from '../../i18n/use-translation';
import { DuotonePanel } from './DuotonePanel';
import { ImageCropSection } from './ImageCropSection';
import { CARD, HEADING } from './inspector-pane-constants';
import { InspectorFields } from './InspectorFields';

// ---------------------------------------------------------------------------
// Props
// ---------------------------------------------------------------------------

interface ImagePropertiesPanelProps {
	selectedElement: PptxImageLikeElement;
	canEdit: boolean;
	onUpdateElement: (updates: Partial<PptxElement>) => void;
}

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

export function ImagePropertiesPanel({
	selectedElement,
	canEdit,
	onUpdateElement,
}: ImagePropertiesPanelProps): React.ReactElement {
	const { t } = useTranslation();

	return (
		<div className={CARD}>
			<div className={HEADING}>{t('pptx.image.properties')}</div>
			<div className='space-y-2'>
				{/* Alt text, adjustments and crop: shared descriptors, one renderer. */}
				<div className='grid grid-cols-2 gap-1.5 text-[11px]'>
					<InspectorFields
						sections={ELEMENT_ALT_TEXT_SECTIONS}
						target={selectedElement as PptxElement}
						disabled={!canEdit}
						onPatch={onUpdateElement}
					/>
					<InspectorFields
						sections={PICTURE_ADJUSTMENT_SECTIONS}
						target={selectedElement as PptxElement}
						disabled={!canEdit}
						onPatch={onUpdateElement}
					/>
				</div>

				<ImageCropSection
					selectedElement={selectedElement}
					canEdit={canEdit}
					onUpdateElement={onUpdateElement}
				/>

				{/* Colour wash, recolour and duotone: shared descriptors, one renderer. */}
				<div className='grid grid-cols-2 gap-1.5 text-[11px]'>
					<InspectorFields
						sections={PICTURE_EFFECT_SECTIONS}
						target={selectedElement as PptxElement}
						disabled={!canEdit}
						onPatch={onUpdateElement}
					/>
				</div>
				<DuotonePanel
					selectedElement={selectedElement}
					canEdit={canEdit}
					onUpdateElement={onUpdateElement}
				/>

				<InspectorFields
					sections={PICTURE_ARTISTIC_SECTIONS}
					target={selectedElement as PptxElement}
					disabled={!canEdit}
					onPatch={onUpdateElement}
				/>

				{/* Reset Image */}
				<button
					type='button'
					disabled={!imageResetState(selectedElement, canEdit, true).enabled}
					className='w-full text-[11px] rounded bg-muted hover:bg-accent border border-border px-2 py-1 text-foreground'
					onClick={() => onUpdateElement(imageResetPatch())}
				>
					{t('pptx.image.resetImage')}
				</button>
			</div>
		</div>
	);
}
