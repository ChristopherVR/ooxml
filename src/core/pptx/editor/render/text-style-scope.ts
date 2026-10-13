/**
 * Which `TextStyle` keys belong to the text BODY and which to a RUN.
 *
 * `TextStyle` is one bag for three OOXML levels (`a:bodyPr`, `a:pPr`, `a:rPr`).
 * The block builder (`text-block-style.ts`, `text-body-layout.ts`) reads the
 * body-level keys from `element.textStyle` only; a run's style never carries
 * them anywhere the renderer looks. So when an inline selection is active and
 * a binding routes the WHOLE update through the selected runs, an Increase
 * Indent (`paragraphMarginLeft`) or a vertical-alignment change lands on run
 * styles nobody reads and nothing happens on screen.
 *
 * `splitTextStyleUpdate` partitions an update so a binding can merge the body
 * half into `element.textStyle` regardless of selection and scope only the
 * run half to the selected characters. `reconcileDecorationFlags` keeps the
 * body-level decoration honest after a run-scoped edit, so the block's CSS
 * `text-decoration` cannot keep painting underline over a run that was just
 * un-underlined.
 */

import type { PptxElement, TextSegment, TextStyle } from '../../core/types';
import { hasTextProperties } from '../../core/types';

import type { TextDecorationFlag } from './text-content-style';
import { isContentSegment } from './text-content-style';

/**
 * Keys read from `element.textStyle` only: `a:bodyPr` attributes, the text
 * warp / 3D scene, autofit, and the paragraph geometry the block builder
 * folds in as the body-wide fallback (indent, margins, spacing).
 */
export const BODY_TEXT_STYLE_KEYS: ReadonlySet<keyof TextStyle> = new Set<keyof TextStyle>([
	// autofit (`a:normAutofit` / `a:spAutoFit`)
	'autoFit',
	'autoFitMode',
	'autoFitFontScale',
	'autoFitLineSpacingReduction',
	// `a:bodyPr` anchoring, direction, rotation
	'vAlign',
	'anchorCenter',
	'textDirection',
	'textBodyRotation',
	'upright',
	// `a:bodyPr` columns, overflow, wrap
	'columnCount',
	'columnSpacing',
	'rtlColumns',
	'hOverflow',
	'vertOverflow',
	'textWrap',
	// `a:bodyPr` insets
	'bodyInsetLeft',
	'bodyInsetTop',
	'bodyInsetRight',
	'bodyInsetBottom',
	// paragraph geometry the block applies body-wide
	'paragraphMarginLeft',
	'paragraphMarginRight',
	'paragraphIndent',
	'paragraphSpacingBefore',
	'paragraphSpacingAfter',
	'lineSpacing',
	'lineSpacingExactPt',
	'tabStops',
	'defaultTabSize',
	'eaLineBreak',
	'latinLineBreak',
	'fontAlignment',
	'hangingPunctuation',
	// `a:bodyPr` flags
	'spaceFirstLastParagraph',
	'fromWordArt',
	'forceAntiAlias',
	'compatibleLineSpacing',
	// warp and 3D scene on the body
	'textWarpPreset',
	'textWarpAdj',
	'textWarpAdj2',
	'text3d',
	'textBodyScene3d',
	'textBodyScene3dXml',
	'flatText',
	'bodyPropertiesExtLstXml',
]);

/** Whether `key` is read from the element body rather than a run. */
export function isBodyTextStyleKey(key: string): boolean {
	return BODY_TEXT_STYLE_KEYS.has(key as keyof TextStyle);
}

/** An update partitioned by the level the renderer reads each key from. */
export interface TextStyleUpdateSplit {
	/** Keys to merge into `element.textStyle` regardless of any selection. */
	body: Partial<TextStyle>;
	/** Keys to scope to the selected runs (or every run without a selection). */
	run: Partial<TextStyle>;
}

/** Partition `updates` into its body-level and run-level halves. Keys set to `undefined` (clears) are kept. */
export function splitTextStyleUpdate(updates: Partial<TextStyle>): TextStyleUpdateSplit {
	const body: Record<string, unknown> = {};
	const run: Record<string, unknown> = {};
	for (const [key, value] of Object.entries(updates)) {
		(isBodyTextStyleKey(key) ? body : run)[key] = value;
	}
	return { body: body as Partial<TextStyle>, run: run as Partial<TextStyle> };
}

/**
 * Recompute the body-level value of each `flag` from the runs after a
 * run-scoped edit: `true` only when every content run is effectively on;
 * otherwise the body flag is dropped, and any run that was only inheriting it
 * gets it written explicitly so its own look does not change.
 */
export function reconcileDecorationFlags(
	segments: readonly TextSegment[],
	textStyle: TextStyle | undefined,
	flags: readonly TextDecorationFlag[],
): { textStyle: TextStyle; textSegments: TextSegment[] } {
	const nextStyle: TextStyle = { ...(textStyle ?? {}) };
	let nextSegments: TextSegment[] = [...segments];
	for (const flag of flags) {
		const body = nextStyle[flag] ?? false;
		const content = nextSegments.filter(isContentSegment);
		if (content.length === 0) {
			continue;
		}
		if (content.every((segment) => segment.style?.[flag] ?? body)) {
			nextStyle[flag] = true;
			continue;
		}
		if (!body) {
			continue;
		}
		delete nextStyle[flag];
		nextSegments = nextSegments.map((segment) =>
			isContentSegment(segment) && segment.style?.[flag] === undefined
				? { ...segment, style: { ...segment.style, [flag]: true } }
				: segment,
		);
	}
	return { textStyle: nextStyle, textSegments: nextSegments };
}

/** `a:lnSpc` holds one of a percentage or an exact size, never both. */
const LINE_SPACING_TWIN: Partial<Record<keyof TextStyle, keyof TextStyle>> = {
	lineSpacing: 'lineSpacingExactPt',
	lineSpacingExactPt: 'lineSpacing',
};

/**
 * Clear the other line-spacing mode when an update sets one. Without this a
 * multiplier applied over an authored `a:spcPts` keeps the exact size, which
 * the renderer prefers, so the edit shows nothing.
 */
export function withExclusiveLineSpacing(updates: Partial<TextStyle>): Partial<TextStyle> {
	const next: Partial<TextStyle> = { ...updates };
	for (const key of Object.keys(updates) as (keyof TextStyle)[]) {
		const twin = LINE_SPACING_TWIN[key];
		if (twin && updates[key] !== undefined && !(twin in updates)) {
			next[twin] = undefined;
		}
	}
	return next;
}

/**
 * Apply a body-wide paragraph update over each paragraph's own `a:pPr`.
 *
 * A paragraph's authored properties (`segment.paragraphProperties`) win over
 * `element.textStyle` both on screen and on save, so an element-wide line
 * spacing or indent change is invisible on any paragraph that declares its
 * own. Only body-level keys the paragraph already declares (or whose
 * line-spacing twin it declares) are rewritten, so nothing new is pinned onto
 * its `a:pPr` and run-level keys such as `align` keep their old behaviour.
 * Pass `updates` through {@link withExclusiveLineSpacing} first.
 */
export function overrideParagraphProperties(
	segments: readonly TextSegment[],
	updates: Partial<TextStyle>,
): TextSegment[] {
	const keys = (Object.keys(updates) as (keyof TextStyle)[]).filter(isBodyTextStyleKey);
	return segments.map((segment) => {
		const own = segment.paragraphProperties;
		const declared = own
			? keys.filter((key) => {
					const twin = LINE_SPACING_TWIN[key];
					return key in own || (twin !== undefined && twin in own);
				})
			: [];
		if (!own || declared.length === 0) {
			return segment;
		}
		const next: TextStyle = { ...own };
		for (const key of declared) {
			if (updates[key] === undefined) {
				delete next[key];
			} else {
				Object.assign(next, { [key]: updates[key] });
			}
		}
		return { ...segment, paragraphProperties: next };
	});
}

/** The parts of a text element below its body style that a paragraph update rewrites. */
export interface ParagraphLevelPatch {
	textSegments?: TextSegment[];
	paragraphIndents?: { marginLeft?: number; indent?: number }[];
}

/**
 * The per-paragraph indents with an element-wide indent applied. The renderer
 * reads these instead of `textStyle` once a body has any
 * (`text-block-style.ts`), so without this an indent change only shows after
 * a save and reload. Empty when there is nothing to rewrite.
 */
export function overrideParagraphIndents(
	element: PptxElement,
	updates: Partial<TextStyle>,
): Pick<ParagraphLevelPatch, 'paragraphIndents'> {
	const { paragraphMarginLeft, paragraphIndent } = updates;
	if (
		!hasTextProperties(element) ||
		!element.paragraphIndents ||
		(paragraphMarginLeft === undefined && paragraphIndent === undefined)
	) {
		return {};
	}
	return {
		paragraphIndents: element.paragraphIndents.map((entry) => ({
			...entry,
			...(paragraphMarginLeft === undefined ? {} : { marginLeft: paragraphMarginLeft }),
			...(paragraphIndent === undefined ? {} : { indent: paragraphIndent }),
		})),
	};
}

/**
 * The paragraph-level half of an element-wide text-style patch: the segments
 * whose own `a:pPr` had to change and the rewritten indents. Empty when
 * nothing below the body changes, so a plain body edit does not resend every
 * segment. Pass `updates` through {@link withExclusiveLineSpacing} first.
 */
export function paragraphOverridePatch(
	element: PptxElement,
	updates: Partial<TextStyle>,
): ParagraphLevelPatch {
	const current = hasTextProperties(element) ? element.textSegments : undefined;
	const overridden = current && overrideParagraphProperties(current, updates);
	const changed = overridden?.some((segment, index) => segment !== current?.[index]);
	return {
		...(changed ? { textSegments: overridden } : {}),
		...overrideParagraphIndents(element, updates),
	};
}
