// Formatting criteria and replacement patches for Word's Find & Replace (Format > Font, Paragraph,
// Style). Matching resolves inherited style and theme properties without mutating the document.
import type { StHighlightColor } from './generated/wml-simple-types';
import type {
	Paragraph,
	ParagraphFormatting,
	ParagraphStyleCatalog,
	RunStyleCatalog,
	TextRun,
	ThemeCatalog,
} from './model';
import type { ParagraphAlign } from './paragraph-alignment';
import { resolveParagraphFormatting } from './paragraph-styles';
import { resolveRunFormatting } from './run-formatting';
import { defaultStyleId } from './style-inheritance';
import { resolveThemeColorReference } from './theme-color';

/** Run, paragraph and style criteria for formatted Find (Word's Format > Font, Paragraph, Style). */
export interface FindFormatCriteria {
	readonly fontFamily?: string;
	readonly fontSizePt?: number;
	readonly bold?: boolean;
	readonly italic?: boolean;
	readonly underline?: boolean;
	readonly color?: string;
	readonly highlight?: StHighlightColor;
	readonly alignment?: ParagraphAlign;
	readonly styleId?: string;
}

/** Formatting changes applied to matched runs and paragraphs on Replace / Replace All. */
export interface ReplaceFormatPatch {
	readonly fontFamily?: string;
	readonly fontSizePt?: number;
	readonly bold?: boolean;
	readonly italic?: boolean;
	readonly underline?: boolean;
	/** Set to `#RRGGBB` (or `RRGGBB`) to apply a color, or `null` to clear direct/theme color. */
	readonly color?: string | null;
	/** Set to a highlight token to apply it, or `'none'` / `null` to clear highlight. */
	readonly highlight?: StHighlightColor | null;
	readonly alignment?: ParagraphAlign;
	/** Paragraph (or character, if defined as such in `runStyles`) style ID, or `null` to clear. */
	readonly styleId?: string | null;
	/** Explicit character style ID (`w:rStyle`) on replaced runs, or `null` to clear. */
	readonly characterStyleId?: string | null;
}

/** Optional style and theme catalogs used to resolve inherited formatting during search/replace. */
export interface FindFormatContext {
	readonly paragraphStyles?: ParagraphStyleCatalog;
	readonly runStyles?: RunStyleCatalog;
	readonly theme?: ThemeCatalog;
}

/** Normalizes a hex color string to `#RRGGBB` uppercase for comparison and storage. */
export function normalizeFormatColor(color: string): string {
	const trimmed = color.trim();
	const hex = trimmed.startsWith('#') ? trimmed.slice(1) : trimmed;
	return `#${hex.toUpperCase()}`;
}

/** Whether `criteria` specifies at least one run, paragraph or style constraint. */
export function hasFormatCriteria(criteria: FindFormatCriteria | undefined): boolean {
	if (!criteria) return false;
	return (
		criteria.fontFamily !== undefined ||
		criteria.fontSizePt !== undefined ||
		criteria.bold !== undefined ||
		criteria.italic !== undefined ||
		criteria.underline !== undefined ||
		criteria.color !== undefined ||
		criteria.highlight !== undefined ||
		criteria.alignment !== undefined ||
		criteria.styleId !== undefined
	);
}

/** Whether `patch` specifies at least one formatting change to apply on replacement. */
export function hasReplaceFormat(patch: ReplaceFormatPatch | undefined): boolean {
	if (!patch) return false;
	return (
		patch.fontFamily !== undefined ||
		patch.fontSizePt !== undefined ||
		patch.bold !== undefined ||
		patch.italic !== undefined ||
		patch.underline !== undefined ||
		patch.color !== undefined ||
		patch.highlight !== undefined ||
		patch.alignment !== undefined ||
		patch.styleId !== undefined ||
		patch.characterStyleId !== undefined
	);
}

function effectiveParagraphStyleId(
	paragraph: Paragraph,
	context: FindFormatContext | undefined,
): string | undefined {
	if (paragraph.style !== undefined) return paragraph.style;
	if (!context?.paragraphStyles) return undefined;
	return defaultStyleId(context.paragraphStyles.styles);
}

function resolvedParagraph(
	paragraph: Paragraph,
	context: FindFormatContext | undefined,
): ParagraphFormatting {
	if (context?.paragraphStyles) {
		return resolveParagraphFormatting(paragraph, context.paragraphStyles);
	}
	return paragraph;
}

function effectiveParagraphAlign(formatting: ParagraphFormatting): ParagraphAlign {
	if (formatting.align !== undefined) return formatting.align;
	return formatting.direction === 'rtl' ? 'right' : 'left';
}

/**
 * Checks whether a paragraph satisfies paragraph-level criteria (`alignment` and, when no run
 * character style could match it, `styleId`).
 */
export function matchesParagraphFormat(
	paragraph: Paragraph,
	criteria: FindFormatCriteria,
	context?: FindFormatContext,
): boolean {
	if (criteria.alignment !== undefined) {
		const align = effectiveParagraphAlign(resolvedParagraph(paragraph, context));
		if (align !== criteria.alignment) return false;
	}
	if (criteria.styleId !== undefined) {
		const isCharStyle = context?.runStyles?.styles[criteria.styleId]?.type === 'character';
		if (!isCharStyle && effectiveParagraphStyleId(paragraph, context) !== criteria.styleId) {
			return false;
		}
	}
	return true;
}

/**
 * Checks whether a single `TextRun` inside `paragraph` satisfies every property specified in
 * `criteria`, resolving inherited paragraph/character styles and theme colors when catalogs are
 * provided in `context`.
 */
export function matchesFormatCriteria(
	run: TextRun,
	paragraph: Paragraph,
	criteria: FindFormatCriteria,
	context?: FindFormatContext,
): boolean {
	if (criteria.alignment !== undefined) {
		const align = effectiveParagraphAlign(resolvedParagraph(paragraph, context));
		if (align !== criteria.alignment) return false;
	}
	if (criteria.styleId !== undefined) {
		const paraStyle = effectiveParagraphStyleId(paragraph, context);
		if (paraStyle !== criteria.styleId && run.style !== criteria.styleId) return false;
	}

	const needsRun =
		criteria.fontFamily !== undefined ||
		criteria.fontSizePt !== undefined ||
		criteria.bold !== undefined ||
		criteria.italic !== undefined ||
		criteria.underline !== undefined ||
		criteria.color !== undefined ||
		criteria.highlight !== undefined;
	if (!needsRun) return true;

	const resolved = resolveRunFormatting(run, {
		paragraphStyleId: paragraph.style,
		paragraphCatalog: context?.paragraphStyles,
		runCatalog: context?.runStyles,
	});

	if (criteria.bold !== undefined && Boolean(resolved.bold) !== criteria.bold) return false;
	if (criteria.italic !== undefined && Boolean(resolved.italic) !== criteria.italic) return false;
	if (criteria.underline !== undefined) {
		const hasUnderline =
			resolved.underline === true ||
			(resolved.underlineStyle !== undefined && resolved.underlineStyle !== 'none');
		if (hasUnderline !== criteria.underline) return false;
	}
	if (criteria.fontSizePt !== undefined && resolved.fontSize !== criteria.fontSizePt) return false;
	if (criteria.fontFamily !== undefined) {
		const actual = resolved.fontFamily?.trim().toLowerCase();
		if (actual !== criteria.fontFamily.trim().toLowerCase()) return false;
	}
	if (criteria.color !== undefined) {
		const actualColor =
			resolved.color ??
			(resolved.colorTheme && context?.theme
				? resolveThemeColorReference(resolved.colorTheme, context.theme)
				: undefined);
		if (!actualColor) return false;
		if (normalizeFormatColor(actualColor) !== normalizeFormatColor(criteria.color)) return false;
	}
	if (criteria.highlight !== undefined) {
		if (criteria.highlight === 'none') {
			if (resolved.highlight !== undefined && resolved.highlight !== 'none') return false;
		} else if (resolved.highlight !== criteria.highlight) {
			return false;
		}
	}
	return true;
}

/** Applies run-level fields of `patch` to a `TextRun` without losing untouched properties. */
export function applyRunFormatPatch(
	run: TextRun,
	patch: ReplaceFormatPatch,
	context?: FindFormatContext,
): TextRun {
	const next: TextRun = { ...run };
	if (patch.fontFamily !== undefined) {
		next.fontFamily = patch.fontFamily;
		delete next.fontTheme;
	}
	if (patch.fontSizePt !== undefined) next.fontSize = patch.fontSizePt;
	if (patch.bold !== undefined) next.bold = patch.bold;
	if (patch.italic !== undefined) next.italic = patch.italic;
	if (patch.underline !== undefined) {
		next.underline = patch.underline;
		if (!patch.underline) {
			delete next.underlineStyle;
			delete next.underlineColor;
		}
	}
	if (patch.color !== undefined) {
		if (patch.color === null) {
			delete next.color;
			delete next.colorTheme;
		} else {
			next.color = normalizeFormatColor(patch.color);
			delete next.colorTheme;
		}
	}
	if (patch.highlight !== undefined) {
		if (patch.highlight === null || patch.highlight === 'none') delete next.highlight;
		else next.highlight = patch.highlight;
	}
	if (patch.characterStyleId !== undefined) {
		if (patch.characterStyleId === null) delete next.style;
		else next.style = patch.characterStyleId;
	} else if (
		typeof patch.styleId === 'string' &&
		context?.runStyles?.styles[patch.styleId]?.type === 'character'
	) {
		next.style = patch.styleId;
	}
	return next;
}

/** Applies paragraph-level fields of `patch` to a `Paragraph` without losing untouched properties. */
export function applyParagraphFormatPatch(
	paragraph: Paragraph,
	patch: ReplaceFormatPatch,
	context?: FindFormatContext,
): Paragraph {
	const next: Paragraph = { ...paragraph };
	if (patch.alignment !== undefined) {
		next.align = patch.alignment;
		delete next.justification;
	}
	if (patch.styleId !== undefined) {
		if (patch.styleId === null) {
			delete next.style;
		} else if (context?.runStyles?.styles[patch.styleId]?.type !== 'character') {
			next.style = patch.styleId;
		}
	}
	return next;
}
