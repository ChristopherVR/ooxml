import {
	resolveShapeFillView,
	SHAPE_FILL_MODE_SECTIONS,
	SHAPE_FILL_OPTIONS_SECTIONS,
	shapeFillChange,
	shapeFillPatch,
} from 'ooxml-ui/pptx';
import type { ShapeFillAction } from 'ooxml-ui/pptx';
import type { PptxThemeColorRef, ShapeStyle } from 'pptx-viewer-core';
import React from 'react';

import { useTranslation } from '../../i18n/use-translation';
import { COL2, LBL, SEL } from './FillStrokeHelpers';
import { ColorPickerRow, GradientStopRow } from './FillStrokeSubComponents';
import { InspectorFields } from './InspectorFields';

export interface ShapeFillSectionProps {
	style: ShapeStyle | undefined;
	/** Lines have no area to fill: the colour pickers are disabled. */
	isLine: boolean;
	onUpdateShapeStyle: (updates: Partial<ShapeStyle>) => void;
	onSetFillColor: (color: string, ref?: PptxThemeColorRef) => void;
}

/**
 * The Format pane's Fill section: the shared fill descriptors drawn by
 * {@link InspectorFields} (kind, transparency, gradient type and angle,
 * picture tiling), and the compound controls `resolveShapeFillView` asks for
 * (solid colour, gradient stops, pattern swatches and colours, picture
 * source), whose events become the shared `shapeFillPatch` actions. Every
 * decision is in `ooxml-ui/pptx`; this component only maps it onto JSX.
 */
export function ShapeFillSection({
	style,
	isLine,
	onUpdateShapeStyle,
	onSetFillColor,
}: ShapeFillSectionProps): React.ReactElement {
	const { t } = useTranslation();
	const view = resolveShapeFillView(style);
	const act = (action: ShapeFillAction): void => {
		const patch = shapeFillPatch(style, action);
		if (patch) {
			onUpdateShapeStyle(patch);
		}
	};
	const styleChange =
		(type: 'patternForeground' | 'patternBackground') =>
		(hex: string, current: ShapeStyle): Partial<ShapeStyle> =>
			shapeFillPatch(current, { type, hex }) ?? {};

	return (
		<>
			<InspectorFields
				sections={SHAPE_FILL_MODE_SECTIONS}
				target={style}
				onPatch={onUpdateShapeStyle}
			/>

			{view.color && (
				<ColorPickerRow
					label={t('pptx.fill.fill')}
					prefix='fill'
					value={view.color.value}
					selectedRef={view.color.colorRef}
					disabled={isLine}
					onChange={onSetFillColor}
					styleChange={(hex) => shapeFillChange(hex)}
				/>
			)}

			<InspectorFields
				sections={SHAPE_FILL_OPTIONS_SECTIONS}
				target={style}
				onPatch={onUpdateShapeStyle}
			/>

			{view.stops && (
				<div className={`${COL2} space-y-1.5 rounded border border-border p-2`}>
					<div className='text-[11px] uppercase tracking-wide text-muted-foreground'>
						{t('pptx.fillAdvanced.gradientStops')}
					</div>
					{view.stops.map((stop) => (
						<GradientStopRow key={`gradient-stop-${stop.index}`} stop={stop} onAction={act} />
					))}
					<button
						type='button'
						className='inline-flex items-center justify-center rounded bg-muted px-2 py-1 text-[11px] hover:bg-accent'
						onClick={() => act({ type: 'addStop' })}
					>
						{t('pptx.gradient.addStop')}
					</button>
				</div>
			)}

			{view.pattern && (
				<>
					<div className={`${COL2} flex flex-col gap-2`}>
						<span className={LBL}>{t('pptx.table.patternPreset')}</span>
						<div className='grid grid-cols-8 gap-1 max-h-48 overflow-y-auto rounded border border-border p-2'>
							{view.pattern.swatches.map((swatch) => (
								<button
									key={swatch.value}
									type='button'
									title={t(swatch.labelKey)}
									aria-label={t(swatch.labelKey)}
									aria-pressed={swatch.selected}
									data-testid='fill-pattern-swatch'
									className={`h-8 w-8 rounded border transition-all ${
										swatch.selected
											? 'border-primary border-2 ring-2 ring-primary/20'
											: 'border-border hover:border-primary/50'
									}`}
									onClick={() => act({ type: 'patternPreset', preset: swatch.value })}
								>
									{swatch.backgroundImage && (
										<div
											className='w-full h-full rounded-sm'
											style={{
												backgroundImage: swatch.backgroundImage,
												backgroundRepeat: 'repeat',
												backgroundSize: '8px 8px',
											}}
										/>
									)}
								</button>
							))}
						</div>
					</div>
					<ColorPickerRow
						label={t('pptx.fillAdvanced.foregroundColor')}
						prefix='pattern-fg'
						value={view.pattern.foreground}
						selectedRef={view.pattern.foregroundRef}
						disabled={isLine}
						onChange={(hex, ref) => act({ type: 'patternForeground', hex, ref })}
						styleChange={styleChange('patternForeground')}
					/>
					<ColorPickerRow
						label={t('pptx.fillAdvanced.backgroundColor')}
						prefix='pattern-bg'
						value={view.pattern.background}
						disabled={isLine}
						onChange={(hex) => act({ type: 'patternBackground', hex })}
						styleChange={styleChange('patternBackground')}
					/>
				</>
			)}

			{view.image && (
				<div className={`${COL2} space-y-2`}>
					<label className='flex flex-col gap-1'>
						<span className={LBL}>{t('pptx.fillAdvanced.imageUrl')}</span>
						<input
							type='text'
							placeholder='https://example.com/image.png'
							value={view.image.url}
							onChange={(e) => act({ type: 'imageUrl', url: e.target.value })}
							className={`${SEL} text-xs`}
						/>
					</label>
					<label className='flex flex-col gap-1'>
						<span className={LBL}>{t('pptx.fillAdvanced.orChooseFile')}</span>
						<input
							type='file'
							accept='image/*'
							onChange={(e) => {
								const file = e.target.files?.[0];
								if (!file) {
									return;
								}
								const reader = new FileReader();
								reader.onload = () => act({ type: 'imageUrl', url: String(reader.result) });
								reader.readAsDataURL(file);
							}}
							className='text-xs text-foreground file:bg-accent file:border-0 file:rounded file:px-2 file:py-1 file:text-xs file:text-foreground file:mr-2'
						/>
					</label>
					{view.image.url && (
						<div className='mt-1 rounded border border-border overflow-hidden h-16'>
							<img
								src={view.image.url}
								alt={t('pptx.fillAdvanced.fillPreviewAlt')}
								className='w-full h-full object-cover'
							/>
						</div>
					)}
				</div>
			)}
		</>
	);
}
