import { DUOTONE_PRESETS, pictureDuotonePresetPatch } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';
import React from 'react';

import { useTranslation } from '../../i18n/use-translation';

interface DuotonePanelProps {
	selectedElement: PptxElement;
	canEdit: boolean;
	onUpdateElement: (updates: Partial<PptxElement>) => void;
}

/**
 * The duotone preset swatches: the native half of the picture's Duotone
 * effect. The toggle and its two colours are the shared
 * `PICTURE_DUOTONE_SECTION`; a swatch applies the shared
 * `pictureDuotonePresetPatch`.
 */
export function DuotonePanel({
	selectedElement,
	canEdit,
	onUpdateElement,
}: DuotonePanelProps): React.ReactElement {
	const { t } = useTranslation();
	return (
		<div className='space-y-1 text-[11px]'>
			<div className='text-muted-foreground text-[10px]'>{t('pptx.image.duotonePresets')}</div>
			<div className='grid grid-cols-4 gap-1'>
				{DUOTONE_PRESETS.map((preset) => (
					<button
						key={preset.labelKey}
						type='button'
						disabled={!canEdit}
						data-testid='picture-duotone-preset'
						className='flex flex-col items-center gap-0.5 rounded border border-border hover:bg-accent p-1 transition-colors'
						title={t(preset.labelKey)}
						onClick={() => onUpdateElement(pictureDuotonePresetPatch(selectedElement, preset))}
					>
						<div className='flex h-4 w-full rounded overflow-hidden'>
							<div className='flex-1' style={{ backgroundColor: preset.shadow }} />
							<div className='flex-1' style={{ backgroundColor: preset.highlight }} />
						</div>
						<span className='text-[8px] text-muted-foreground truncate w-full text-center'>
							{t(preset.labelKey)}
						</span>
					</button>
				))}
			</div>
		</div>
	);
}
