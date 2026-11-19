/**
 * Accessibility checker for PPTX presentations.
 *
 * Scans slides for common accessibility issues: missing alt text,
 * missing slide titles, low contrast text, complex merged tables,
 * duplicate titles, and blank slides.
 *
 * Implements checks aligned with WCAG 2.1 AA guidelines and
 * PowerPoint's built-in accessibility checker.
 *
 * @module utils/accessibility-checker
 */

import type { PptxElement } from '../types/elements';
import type { PptxData, PptxSlide } from '../types/presentation';
import type { PptxTableData } from '../types/table';
import { at } from '../../../geometry/indexed';
import { contrastRatio, parseHexColor, relativeLuminance } from '../../../color/contrast';

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

export type AccessibilityIssueType =
	| 'missingAltText'
	| 'missingSlideTitle'
	| 'lowContrast'
	| 'complexTable'
	| 'duplicateTitle'
	| 'blankSlide';

export type AccessibilityIssueSeverity = 'error' | 'warning' | 'tip';

export interface AccessibilityIssue {
	type: AccessibilityIssueType;
	severity: AccessibilityIssueSeverity;
	slideIndex: number;
	elementId?: string;
	message: string;
	suggestion: string;
}

export interface AccessibilityCheckOptions {
	/** Minimum WCAG contrast ratio (default 4.5 for AA normal text). */
	minContrastRatio?: number;
	/** Skip contrast checks (useful when background can't be resolved). */
	skipContrast?: boolean;
	/** Skip blank slide checks. */
	skipBlankSlide?: boolean;
}

// ---------------------------------------------------------------------------
// Colour helpers (WCAG 2.x, shared in ooxml-core/color)
// ---------------------------------------------------------------------------

export { parseHexColor, relativeLuminance };

/** WCAG contrast ratio between two hex colours, in [1, 21]; an unparseable colour counts as 21. */
export function computeContrastRatio(fg: string, bg: string): number {
	return contrastRatio(fg, bg) ?? 21;
}

// ---------------------------------------------------------------------------
// Element helpers
// ---------------------------------------------------------------------------

const IMAGE_LIKE_TYPES = new Set(['image', 'picture']);
const VISUAL_CONTENT_TYPES = new Set(['image', 'picture', 'chart', 'smartArt', 'media']);

function getElementText(el: PptxElement): string {
	if ('text' in el && typeof el.text === 'string') {
		return el.text;
	}
	if ('textSegments' in el && Array.isArray(el.textSegments)) {
		return el.textSegments.map((s: { text?: string }) => s.text || '').join('');
	}
	return '';
}

function getAltText(el: PptxElement): string | undefined {
	if ('altText' in el) {
		return el.altText as string | undefined;
	}
	return undefined;
}

function getTextColor(el: PptxElement): string | undefined {
	if ('textStyle' in el && el.textStyle) {
		return (el.textStyle as { color?: string }).color;
	}
	if ('textSegments' in el && Array.isArray(el.textSegments) && el.textSegments.length > 0) {
		const firstStyle = (el.textSegments as Array<{ style?: { color?: string } }>)[0]?.style;
		return firstStyle?.color;
	}
	return undefined;
}

function getBackgroundColor(el: PptxElement): string | undefined {
	if ('shapeStyle' in el && el.shapeStyle) {
		return (el.shapeStyle as { fillColor?: string }).fillColor;
	}
	return undefined;
}

function isTitleElement(el: PptxElement): boolean {
	const id = el.id.toLowerCase();
	return id.includes('title') || id.startsWith('ctr') || id.includes('ctrtitle');
}

function isVisibleElement(el: PptxElement): boolean {
	if ('hidden' in el && el.hidden) {
		return false;
	}
	if ('opacity' in el && typeof el.opacity === 'number' && el.opacity <= 0) {
		return false;
	}
	return true;
}

// ---------------------------------------------------------------------------
// Individual checks
// ---------------------------------------------------------------------------

export function checkMissingAltText(slide: PptxSlide, slideIndex: number): AccessibilityIssue[] {
	const issues: AccessibilityIssue[] = [];
	for (const el of slide.elements) {
		if (!isVisibleElement(el)) {
			continue;
		}
		if (!VISUAL_CONTENT_TYPES.has(el.type)) {
			continue;
		}

		const alt = getAltText(el);
		if (!alt || alt.trim() === '') {
			const label = IMAGE_LIKE_TYPES.has(el.type)
				? 'Image'
				: el.type === 'chart'
					? 'Chart'
					: el.type === 'smartArt'
						? 'SmartArt'
						: 'Media';
			issues.push({
				type: 'missingAltText',
				severity: 'error',
				slideIndex,
				elementId: el.id,
				message: `${label} is missing alternative text.`,
				suggestion: `Add a description to the ${label.toLowerCase()} so screen readers can describe it.`,
			});
		}
	}
	return issues;
}

export function checkMissingSlideTitle(slide: PptxSlide, slideIndex: number): AccessibilityIssue[] {
	const titleElement = slide.elements.find((el) => isTitleElement(el) && isVisibleElement(el));
	if (!titleElement) {
		return [
			{
				type: 'missingSlideTitle',
				severity: 'warning',
				slideIndex,
				message: 'Slide is missing a title.',
				suggestion: 'Add a title element so screen readers can navigate between slides.',
			},
		];
	}
	const text = getElementText(titleElement).trim();
	if (text === '') {
		return [
			{
				type: 'missingSlideTitle',
				severity: 'warning',
				slideIndex,
				elementId: titleElement.id,
				message: 'Slide title is empty.',
				suggestion: 'Add text to the title element.',
			},
		];
	}
	return [];
}

export function checkLowContrast(
	slide: PptxSlide,
	slideIndex: number,
	minRatio: number,
	slideBg?: string,
): AccessibilityIssue[] {
	const issues: AccessibilityIssue[] = [];
	const defaultBg = slideBg || '#FFFFFF';

	for (const el of slide.elements) {
		if (!isVisibleElement(el)) {
			continue;
		}
		const text = getElementText(el);
		if (!text.trim()) {
			continue;
		}

		const fg = getTextColor(el);
		if (!fg) {
			continue;
		} // can't determine colour → skip

		const bg = getBackgroundColor(el) || defaultBg;
		const ratio = computeContrastRatio(fg, bg);

		if (ratio < minRatio) {
			issues.push({
				type: 'lowContrast',
				severity: 'warning',
				slideIndex,
				elementId: el.id,
				message: `Text contrast ratio is ${ratio.toFixed(1)}:1 (minimum ${minRatio}:1).`,
				suggestion: 'Increase contrast between text colour and background.',
			});
		}
	}
	return issues;
}

export function checkComplexTables(slide: PptxSlide, slideIndex: number): AccessibilityIssue[] {
	const issues: AccessibilityIssue[] = [];
	for (const el of slide.elements) {
		if (el.type !== 'table') {
			continue;
		}
		const tableData = (el as { tableData?: PptxTableData }).tableData;
		if (!tableData) {
			continue;
		}

		let mergeCount = 0;
		for (const row of tableData.rows) {
			for (const cell of row.cells) {
				if ((cell.gridSpan && cell.gridSpan > 1) || (cell.rowSpan && cell.rowSpan > 1)) {
					mergeCount++;
				}
			}
		}

		if (mergeCount > 2) {
			issues.push({
				type: 'complexTable',
				severity: 'warning',
				slideIndex,
				elementId: el.id,
				message: `Table has ${mergeCount} merged cell regions, which can confuse screen readers.`,
				suggestion: 'Simplify the table structure or split into multiple simpler tables.',
			});
		}
	}
	return issues;
}

export function checkDuplicateTitles(slides: PptxSlide[]): AccessibilityIssue[] {
	const issues: AccessibilityIssue[] = [];
	const titleMap = new Map<string, number[]>();

	for (let i = 0; i < slides.length; i++) {
		const slide = at(slides, i);
		const titleEl = slide.elements.find((el) => isTitleElement(el) && isVisibleElement(el));
		if (!titleEl) {
			continue;
		}
		const text = getElementText(titleEl).trim().toLowerCase();
		if (!text) {
			continue;
		}

		const existing = titleMap.get(text);
		if (existing) {
			existing.push(i);
		} else {
			titleMap.set(text, [i]);
		}
	}

	for (const [title, indices] of titleMap) {
		if (indices.length > 1) {
			for (const idx of indices) {
				issues.push({
					type: 'duplicateTitle',
					severity: 'tip',
					slideIndex: idx,
					message: `Slide title "${title}" is duplicated across ${indices.length} slides.`,
					suggestion: 'Use unique titles to help screen reader users distinguish slides.',
				});
			}
		}
	}
	return issues;
}

export function checkBlankSlide(slide: PptxSlide, slideIndex: number): AccessibilityIssue[] {
	const hasVisibleContent = slide.elements.some((el) => {
		if (!isVisibleElement(el)) {
			return false;
		}
		if (el.type === 'text' || el.type === 'shape') {
			return getElementText(el).trim() !== '';
		}
		return true; // non-text elements count as content
	});

	if (!hasVisibleContent) {
		return [
			{
				type: 'blankSlide',
				severity: 'tip',
				slideIndex,
				message: 'Slide appears to have no visible content.',
				suggestion: 'Add content or remove the blank slide.',
			},
		];
	}
	return [];
}

// ---------------------------------------------------------------------------
// Main entry point
// ---------------------------------------------------------------------------

/**
 * Run all accessibility checks on a presentation.
 *
 * @param data - The parsed presentation data.
 * @param options - Optional configuration for checks.
 * @returns Array of accessibility issues sorted by slide index.
 */
export function checkPresentation(
	data: PptxData,
	options: AccessibilityCheckOptions = {},
): AccessibilityIssue[] {
	const { minContrastRatio = 4.5, skipContrast = false, skipBlankSlide = false } = options;
	const issues: AccessibilityIssue[] = [];

	for (let i = 0; i < data.slides.length; i++) {
		const slide = at(data.slides, i);
		issues.push(...checkMissingAltText(slide, i));
		issues.push(...checkMissingSlideTitle(slide, i));
		if (!skipContrast) {
			issues.push(...checkLowContrast(slide, i, minContrastRatio, slide.backgroundColor));
		}
		issues.push(...checkComplexTables(slide, i));
		if (!skipBlankSlide) {
			issues.push(...checkBlankSlide(slide, i));
		}
	}

	issues.push(...checkDuplicateTitles(data.slides));

	// Sort by slide index, then severity (error > warning > tip)
	const severityOrder: Record<AccessibilityIssueSeverity, number> = {
		error: 0,
		warning: 1,
		tip: 2,
	};
	issues.sort(
		(a, b) => a.slideIndex - b.slideIndex || severityOrder[a.severity] - severityOrder[b.severity],
	);

	return issues;
}
