/**
 * Framework-agnostic stroke/dash utilities for PPTX shapes and connectors.
 */
import { capCompensatedDashArray, dashPatternUnits } from '../../../drawingml/dash-pattern';
import type { StrokeDashType } from '../types';

/**
 * Normalize a raw dash-type string (possibly from OpenXML) into a typed `StrokeDashType`.
 */
export function normalizeStrokeDashType(
	value: StrokeDashType | string | undefined,
): StrokeDashType | undefined {
	const normalized = String(value || '')
		.trim()
		.toLowerCase();
	if (!normalized) {
		return undefined;
	}

	const dashMap: Record<string, StrokeDashType> = {
		solid: 'solid',
		dot: 'dot',
		dash: 'dash',
		lgdash: 'lgDash',
		dashdot: 'dashDot',
		lgdashdot: 'lgDashDot',
		lgdashdotdot: 'lgDashDotDot',
		sysdot: 'sysDot',
		sysdash: 'sysDash',
		sysdashdot: 'sysDashDot',
		sysdashdotdot: 'sysDashDotDot',
		custom: 'custom',
	};
	return dashMap[normalized];
}

/**
 * Map a `StrokeDashType` to the equivalent CSS `borderStyle`.
 */
export function getCssBorderDashStyle(dashType: StrokeDashType | undefined): string | undefined {
	if (!dashType || dashType === 'solid') {
		return 'solid';
	}
	if (dashType === 'dot' || dashType === 'sysDot') {
		return 'dotted';
	}
	return 'dashed';
}

/**
 * Map a `StrokeDashType` to an SVG `stroke-dasharray` value, using PowerPoint's
 * measured lengths (see `ooxml-core/drawingml` `dash-pattern`): the preset
 * multiples of the pen width (`dash` 4 3, `dot` 1 3, `lgDash` 8 3, `sysDot`
 * 1 1, ...), the pen being the stroke width rounded to whole pixels. When
 * `dashType` is `"custom"` and `customDashSegments` are provided (`a:ds`
 * lengths in 1000ths of a percent of the line width) the pattern is built from
 * them. With a `rnd` cap each dash gives up one pen width to its round ends.
 */
export function getSvgStrokeDasharray(
	dashType: StrokeDashType | undefined,
	strokeWidth: number,
	customDashSegments?: Array<{ dash: number; space: number }>,
	lineCap?: 'flat' | 'rnd' | 'sq',
): string | undefined {
	const pattern = dashPatternUnits(dashType, customDashSegments);
	return pattern ? capCompensatedDashArray(pattern, strokeWidth, lineCap) : undefined;
}
