import { cropResetPatch, cropResetState, PICTURE_CROP_SECTIONS } from 'ooxml-ui/pptx';
/**
 * ImageCropSection: the four crop insets (the shared `PICTURE_CROP_SECTIONS`
 * drawn by the generic renderer) plus a reset, inside the image-properties card.
 *
 * WHY it is here at all: cropping had no other entry point in React. The
 * component existed but was only reachable from the unreferenced
 * `ElementProperties`, so the viewer parsed and re-serialised `a:srcRect`
 * faithfully while offering no way to change it. Alt text and "replace image"
 * are deliberately NOT repeated here; `ImagePropertiesPanel` already owns the
 * former and there is no image-picker handler on the live path for the latter.
 */
import type { PptxElement } from 'pptx-viewer-core';
import { isImageLikeElement } from 'pptx-viewer-core';
import React from 'react';

import { useTranslation } from '../../i18n/use-translation';
import { BTN_CLS } from './element-properties-constants';
import { InspectorFields } from './InspectorFields';

interface ImageCropSectionProps {
	selectedElement: PptxElement;
	canEdit: boolean;
	onUpdateElement: (updates: Partial<PptxElement>) => void;
}

export function ImageCropSection({
	selectedElement,
	canEdit,
	onUpdateElement,
}: ImageCropSectionProps): React.ReactElement | null {
	const { t } = useTranslation();
	if (!isImageLikeElement(selectedElement)) {
		return null;
	}
	// G7: `a:picLocks/@noCrop` forbids cropping this specific picture, even on
	// an otherwise-editable deck.
	const croppable = cropResetState(selectedElement, canEdit, true).enabled;

	return (
		<div className='space-y-1.5' data-pptx-image-crop>
			<div className='grid grid-cols-2 gap-1.5 text-[11px]'>
				<InspectorFields
					sections={PICTURE_CROP_SECTIONS}
					target={selectedElement as PptxElement}
					disabled={!croppable}
					onPatch={onUpdateElement}
				/>
			</div>
			<button
				type='button'
				disabled={!croppable}
				className={`${BTN_CLS} w-full`}
				onClick={() => onUpdateElement(cropResetPatch())}
			>
				{t('pptx.image.resetCrop')}
			</button>
		</div>
	);
}
