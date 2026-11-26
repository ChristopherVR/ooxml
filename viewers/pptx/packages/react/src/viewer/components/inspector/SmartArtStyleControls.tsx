import { schemaLabel, SMARTART_STYLE_LABEL_KEYS } from 'ooxml-ui/pptx';
import type { smartArtInspectorColors } from 'ooxml-ui/pptx';
import type { SmartArtStyle } from 'pptx-viewer-core';
import React from 'react';

import { useTranslation } from '../../i18n/use-translation';
import { cn } from '../../utils';
import { WebSelect } from '../WebControls';
import { INPUT } from './inspector-pane-constants';

/**
 * Colour-scheme select + style toggle group for the SmartArt inspector.
 *
 * Lifted out of `SmartArtPropertiesPanel` purely to keep that file inside the
 * per-file line budget; the markup, testids and accessible names are unchanged.
 *
 */
export interface SmartArtStyleControlsProps {
	colors: ReturnType<typeof smartArtInspectorColors>;
	styleOptions: readonly SmartArtStyle[];
	style: SmartArtStyle;
	canEdit: boolean;
	onChangeColorScheme: (scheme: string) => void;
	onChangeStyle: (style: SmartArtStyle) => void;
}

export function SmartArtStyleControls({
	colors,
	styleOptions,
	style,
	canEdit,
	onChangeColorScheme,
	onChangeStyle,
}: SmartArtStyleControlsProps): React.ReactElement {
	const { t } = useTranslation();
	// `schemaLabel` takes a plain `(key) => string`; react-i18next's `t` is an
	// overloaded generic, so narrow it once instead of at every call site.
	const translate = React.useCallback((key: string): string => t(key), [t]);

	return (
		<>
			<label className='flex flex-col gap-1 text-[11px]'>
				<span className='text-muted-foreground'>{t('pptx.smartart.colorScheme')}</span>
				<WebSelect
					disabled={!canEdit}
					data-testid='smartart-color-scheme'
					aria-label={t('pptx.smartart.colorScheme')}
					className={cn(INPUT, 'w-full')}
					value={colors.value}
					onChange={(e) => onChangeColorScheme(e.target.value)}
				>
					{colors.options.map(({ value, label }) => (
						<option key={value} value={value}>
							{label}
						</option>
					))}
				</WebSelect>
			</label>

			<label className='flex flex-col gap-1 text-[11px]'>
				<span className='text-muted-foreground'>{t('pptx.smartart.style')}</span>
				<div className='flex gap-1' role='group' aria-label={t('pptx.smartart.style')}>
					{styleOptions.map((s) => (
						<button
							key={s}
							type='button'
							disabled={!canEdit}
							aria-pressed={style === s}
							className={cn(
								'flex-1 px-2 py-1 text-[10px] rounded border transition-colors',
								style === s
									? 'border-primary bg-primary/20 text-primary'
									: 'border-border text-muted-foreground hover:bg-muted',
							)}
							onClick={() => onChangeStyle(s)}
						>
							{schemaLabel(SMARTART_STYLE_LABEL_KEYS, s, translate)}
						</button>
					))}
				</div>
			</label>
		</>
	);
}
