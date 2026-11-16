import { OFFICE_COLOR_SWATCHES, officeColorSwatchLabel, shapeQuickStylePatch } from 'ooxml-ui/pptx';
import type { ShapeFillAction, ShapeFillStopView } from 'ooxml-ui/pptx';
import type { PptxElement, PptxThemeColorRef, ShapeStyle } from 'pptx-viewer-core';
import { hasShapeProperties } from 'pptx-viewer-core';
import React from 'react';
import { LuPipette } from 'react-icons/lu';

import { useTranslation } from '../../i18n/use-translation';
import { openNativeEyeDropper } from '../../utils';
import { SEL, RNG, SWATCH, DIS, LBL } from './FillStrokeHelpers';
import { LiveColorInput } from './LiveColorInput';
import { useRecentColors } from './RecentColorsContext';
import { RecentColorsRow } from './RecentColorsRow';
import { ThemeColorSwatchGrid } from './ThemeColorSwatchGrid';
import { useRecentColorCommit } from './useRecentColorCommit';

// ---------------------------------------------------------------------------
// ColorPickerRow
// ---------------------------------------------------------------------------

/**
 * Color picker + theme colours + standard colours + recent colors + eyedropper.
 *
 * The recent-colours row is sourced from {@link useRecentColors} (context),
 * not a prop: every caller used to have to thread its own `recentColors`
 * array down from the loaded deck, and one of the two calls (the fill
 * colour's own row) simply never did, so its row silently rendered nothing.
 * Every commit here (native picker, theme swatch, standard swatch, recent
 * swatch, eyedropper) also pushes the resolved hex back into that same shared
 * list; a drag through the native picker shows every shade live but is one
 * undo step and adds only the colour it settles on ({@link LiveColorInput}).
 *
 * Clicking a {@link ThemeColorSwatchGrid} swatch commits BOTH the resolved
 * hex and its `PptxThemeColorRef`, so the colour keeps following the theme
 * after a later theme change; every other commit path (native picker,
 * standard swatch, recent swatch, eyedropper) clears the ref, since none of
 * those carry a theme identity.
 */
export function ColorPickerRow({
	label,
	value,
	selectedRef,
	disabled,
	prefix,
	onChange,
	styleChange,
}: {
	label: string;
	value: string;
	/** The element's current theme ref, if any (highlights the matching theme swatch). */
	selectedRef?: PptxThemeColorRef;
	disabled?: boolean;
	prefix: string;
	onChange: (c: string, ref?: PptxThemeColorRef) => void;
	/**
	 * The shape-style change a native pick of `hex` makes (what `onChange` commits
	 * for it), so a drag through the picker previews and is put back leaf by leaf
	 * ({@link LiveColorInput}). Without it the native picker applies every shade.
	 */
	styleChange?: (hex: string, style: ShapeStyle) => Partial<ShapeStyle>;
}): React.ReactElement {
	const { t } = useTranslation();
	const { pushColor } = useRecentColors();
	const commit = (color: string, ref?: PptxThemeColorRef): void => {
		onChange(color, ref);
		pushColor(color);
	};
	const handleEyedropper = async (): Promise<void> => {
		const color = await openNativeEyeDropper();
		if (color) {
			commit(color);
		}
	};

	return (
		<label className='flex flex-col gap-1'>
			<span className={LBL}>{label}</span>
			<div className='flex items-center gap-1'>
				<LiveColorInput
					value={value}
					disabled={disabled}
					onApply={onChange}
					patch={styleChange && ((el, hex) => shapeStyleChangePatch(el, hex, styleChange))}
					className={`h-8 flex-1 ${SEL} px-1 ${DIS}`}
				/>
				<button
					type='button'
					disabled={disabled}
					// `max-md:h-11 max-md:w-11` (44px) matches MIN_TOUCH_TARGET_PX from
					// pptx-viewer-shared's render/responsive module at Tailwind's `md`
					// (768px) breakpoint; the desktop-mouse `h-8 w-8` (32px) is
					// otherwise correct and left alone above that width.
					className='h-8 w-8 max-md:h-11 max-md:w-11 flex items-center justify-center rounded border border-border bg-muted hover:bg-accent disabled:opacity-40 disabled:cursor-not-allowed transition-colors'
					title={t('pptx.fillStroke.eyedropperTooltip')}
					aria-label={t('pptx.fillStroke.eyedropperTooltip')}
					onClick={(e) => {
						e.preventDefault();
						void handleEyedropper();
					}}
				>
					<LuPipette className='w-3.5 h-3.5' />
				</button>
			</div>
			<ThemeColorSwatchGrid
				disabled={disabled}
				selectedRef={selectedRef}
				selectedHex={value}
				onPick={(c) => commit(c.hex, c.ref)}
			/>
			<div className='mt-1'>
				<div className='text-[10px] text-muted-foreground mb-1'>
					{t('pptx.colorPicker.standardColors')}
				</div>
				<div className='flex flex-wrap gap-1'>
					{OFFICE_COLOR_SWATCHES.map((c) => (
						<button
							key={`${prefix}-standard-${c.hex}`}
							type='button'
							className={`${SWATCH} ${DIS}`}
							style={{ backgroundColor: c.hex }}
							title={`${label} ${officeColorSwatchLabel(c, (key) => t(key))}`}
							aria-label={`${label} ${officeColorSwatchLabel(c, (key) => t(key))}`}
							data-pptx-compact
							disabled={disabled}
							onClick={() => commit(c.hex)}
						/>
					))}
				</div>
			</div>
			<RecentColorsRow disabled={disabled} onCommit={commit} />
		</label>
	);
}

/** The element patch a {@link ColorPickerRow} style change makes for `hex`. */
function shapeStyleChangePatch(
	el: PptxElement,
	hex: string,
	change: (hex: string, style: ShapeStyle) => Partial<ShapeStyle>,
): Partial<PptxElement> | undefined {
	return hasShapeProperties(el)
		? shapeQuickStylePatch(el, change(hex, el.shapeStyle ?? {}))
		: undefined;
}

// ---------------------------------------------------------------------------
// GradientStopRow
// ---------------------------------------------------------------------------

/**
 * One row of the Fill section's gradient stop list, drawn from the shared
 * `ShapeFillStopView`; every control reports a shared `ShapeFillAction`. The
 * theme colour grid sits below the native colour input, same "swatch commits
 * hex + ref, native input clears it" contract as {@link ColorPickerRow}:
 * picking a theme swatch keeps this stop following the deck's theme after a
 * later theme change.
 */
export function GradientStopRow({
	stop,
	onAction,
}: {
	stop: ShapeFillStopView;
	onAction: (action: ShapeFillAction) => void;
}): React.ReactElement {
	const { t } = useTranslation();
	const { pushColor } = useRecentColors();
	const commitRecentColor = useRecentColorCommit();
	const { index } = stop;
	return (
		<div className='space-y-1' data-testid='fill-gradient-stop'>
			<div className='grid grid-cols-[auto,1fr,auto] items-center gap-2'>
				<input
					type='color'
					aria-label={t('pptx.gradient.stops')}
					value={stop.color}
					ref={commitRecentColor}
					onChange={(e) => onAction({ type: 'stopColor', index, hex: e.target.value })}
					className='h-7 w-10 rounded border border-border bg-muted'
				/>
				<input
					type='range'
					aria-label={t('pptx.gradient.position')}
					min={0}
					max={100}
					value={stop.position}
					onChange={(e) =>
						onAction({ type: 'stopPosition', index, position: Number(e.target.value) })
					}
					className={RNG}
				/>
				<button
					type='button'
					disabled={!stop.removable}
					title={t('pptx.gradient.removeStop')}
					className='rounded bg-muted px-2 py-1 text-[11px] hover:bg-accent disabled:opacity-40 disabled:cursor-not-allowed'
					onClick={() => onAction({ type: 'removeStop', index })}
				>
					{t('pptx.comments.remove')}
				</button>
			</div>
			<ThemeColorSwatchGrid
				selectedRef={stop.colorRef}
				selectedHex={stop.color}
				onPick={(c) => {
					onAction({ type: 'stopColor', index, hex: c.hex, ref: c.ref });
					pushColor(c.hex);
				}}
			/>
			<div className='grid grid-cols-[auto,1fr,auto] items-center gap-2 pl-1'>
				<span className='text-[10px] text-muted-foreground w-10 text-center'>
					{t('pptx.strokeEffects.fillOpacity')}
				</span>
				<input
					type='range'
					aria-label={t('pptx.strokeEffects.fillOpacity')}
					min={0}
					max={100}
					value={stop.opacity}
					onChange={(e) =>
						onAction({ type: 'stopOpacity', index, opacity: Number(e.target.value) })
					}
					className={RNG}
				/>
				<span className='text-[10px] text-muted-foreground w-[52px] text-right'>
					{stop.opacity}%
				</span>
			</div>
		</div>
	);
}
