/**
 * The compound controls of the inspector's Fill section, written once: what
 * each binding draws besides the generic fields (`shape-fill-fields.ts`), and
 * the typed actions its events become.
 *
 * {@link resolveShapeFillView} says which blocks the current fill kind shows
 * (the solid colour picker, the gradient stop list, the pattern swatches and
 * colours, the picture source) and with which values; {@link shapeFillPatch}
 * turns a {@link ShapeFillAction} into the `Partial<ShapeStyle>` the binding
 * commits through its usual shape-style route. A binding keeps only its native
 * template: its colour picker with theme grid and recent colours, a stop row,
 * a swatch button.
 *
 * @module render/inspector/shape-fill-actions
 */
import type { PptxThemeColorRef, ShapeStyle } from 'ooxml-core/pptx';
import { normalizeHexColor } from 'ooxml-core/color';

import { getPatternSvg } from '../fill-style';
import { PATTERN_PRESET_OPTIONS } from '../fill-pattern-label-keys';
import { shapeFillChange } from '../ribbon-home-commands';
import {
	DEFAULT_PATTERN_BACKGROUND,
	DEFAULT_PATTERN_PRESET,
	shapeFillColorOf,
	shapeFillModeChange,
	shapeFillModeOf,
} from './shape-fill-fields';
import type { ShapeFillMode, ShapeFillTarget } from './shape-fill-fields';
import {
	addShapeFillStop,
	editShapeFillStop,
	removeShapeFillStop,
	shapeFillStopsOf,
} from './shape-fill-stops';

/** One stop row of the gradient stop list. */
export interface ShapeFillStopView {
	index: number;
	/** `#rrggbb` for the stop's colour input. */
	color: string;
	/** Its theme colour, for the theme grid's highlight. */
	colorRef?: PptxThemeColorRef;
	/** 0-100, rounded. */
	position: number;
	/** Stop opacity, 0-100, rounded. */
	opacity: number;
	/** False while only two stops are left (a gradient needs two). */
	removable: boolean;
}

/** One swatch of the pattern grid. */
export interface ShapeFillPatternSwatch {
	value: string;
	labelKey: string;
	selected: boolean;
	/** CSS `background-image` of the swatch tile in the shape's colours, or `undefined`. */
	backgroundImage?: string;
}

/** What a binding draws for the Fill section besides the generic fields. */
export interface ShapeFillView {
	mode: ShapeFillMode;
	/** The solid colour picker (solid fill only). */
	color?: { value: string; colorRef?: PptxThemeColorRef };
	/** The gradient stop list (gradient fill only). */
	stops?: ShapeFillStopView[];
	/** The pattern swatches and its two colours (pattern fill only). */
	pattern?: {
		swatches: ShapeFillPatternSwatch[];
		foreground: string;
		foregroundRef?: PptxThemeColorRef;
		background: string;
	};
	/** The picture source (picture fill only); `url` is empty when none is set. */
	image?: { url: string };
}

/** What a Fill section control did. */
export type ShapeFillAction =
	| { type: 'mode'; mode: ShapeFillMode }
	| { type: 'color'; hex: string; ref?: PptxThemeColorRef }
	| { type: 'patternPreset'; preset: string }
	| { type: 'patternForeground'; hex: string; ref?: PptxThemeColorRef }
	| { type: 'patternBackground'; hex: string }
	| { type: 'addStop' }
	| { type: 'removeStop'; index: number }
	| { type: 'stopColor'; index: number; hex: string; ref?: PptxThemeColorRef }
	| { type: 'stopPosition'; index: number; position: number }
	| { type: 'stopOpacity'; index: number; opacity: number }
	| { type: 'imageUrl'; url: string };

const percent = (value: number | undefined): number => Math.round((value ?? 1) * 100);

function patternView(style: ShapeFillTarget): NonNullable<ShapeFillView['pattern']> {
	const foreground = shapeFillColorOf(style);
	const background = normalizeHexColor(
		style?.fillPatternBackgroundColor,
		DEFAULT_PATTERN_BACKGROUND,
	);
	const preset = style?.fillPatternPreset || DEFAULT_PATTERN_PRESET;
	const view: NonNullable<ShapeFillView['pattern']> = {
		foreground,
		background,
		swatches: PATTERN_PRESET_OPTIONS.map((option) => {
			const svg = getPatternSvg(option.value, foreground, background);
			const swatch: ShapeFillPatternSwatch = {
				value: option.value,
				labelKey: option.labelKey,
				selected: option.value === preset,
			};
			if (svg) {
				swatch.backgroundImage = `url('data:image/svg+xml;utf8,${encodeURIComponent(svg)}')`;
			}
			return swatch;
		}),
	};
	if (style?.fillColorRef) {
		view.foregroundRef = style.fillColorRef;
	}
	return view;
}

/** The compound controls the current fill kind of `style` shows. */
export function resolveShapeFillView(style: ShapeFillTarget): ShapeFillView {
	const mode = shapeFillModeOf(style);
	const view: ShapeFillView = { mode };
	if (mode === 'solid') {
		view.color = { value: shapeFillColorOf(style) };
		if (style?.fillColorRef) {
			view.color.colorRef = style.fillColorRef;
		}
	} else if (mode === 'gradient') {
		const stops = shapeFillStopsOf(style);
		view.stops = stops.map((stop, index) => {
			const row: ShapeFillStopView = {
				index,
				color: normalizeHexColor(stop.color, '#000000'),
				position: Math.round(stop.position),
				opacity: percent(stop.opacity),
				removable: stops.length > 2,
			};
			if (stop.colorRef) {
				row.colorRef = stop.colorRef;
			}
			return row;
		});
	} else if (mode === 'pattern') {
		view.pattern = patternView(style);
	} else if (mode === 'image') {
		view.image = { url: style?.fillImageUrl ?? '' };
	}
	return view;
}

const clamp = (value: number, lo: number, hi: number): number => Math.max(lo, Math.min(hi, value));

/**
 * The style change `action` makes on `style`, or `undefined` when it changes
 * nothing (an unknown stop, removing one of the last two, a non-numeric box).
 */
export function shapeFillPatch(
	style: ShapeFillTarget,
	action: ShapeFillAction,
): Partial<ShapeStyle> | undefined {
	const leaveBackground: Partial<ShapeStyle> = style?.useBackgroundFill
		? { useBackgroundFill: false }
		: {};
	switch (action.type) {
		case 'mode':
			return shapeFillModeChange(style, action.mode);
		case 'color':
			return { ...shapeFillChange(action.hex, action.ref), ...leaveBackground };
		case 'patternPreset':
			return { ...leaveBackground, fillMode: 'pattern', fillPatternPreset: action.preset };
		case 'patternForeground':
			return {
				...leaveBackground,
				fillMode: 'pattern',
				fillColor: action.hex,
				fillColorRef: action.ref,
			};
		case 'patternBackground':
			return { ...leaveBackground, fillMode: 'pattern', fillPatternBackgroundColor: action.hex };
		case 'addStop':
			return addShapeFillStop(style);
		case 'removeStop':
			return removeShapeFillStop(style, action.index);
		case 'stopColor':
			return editShapeFillStop(style, action.index, { color: action.hex, colorRef: action.ref });
		case 'stopPosition':
			return Number.isFinite(action.position)
				? editShapeFillStop(style, action.index, { position: action.position })
				: undefined;
		case 'stopOpacity':
			return Number.isFinite(action.opacity)
				? editShapeFillStop(style, action.index, { opacity: clamp(action.opacity, 0, 100) / 100 })
				: undefined;
		case 'imageUrl':
			return { ...leaveBackground, fillMode: 'image', fillImageUrl: action.url };
	}
}
