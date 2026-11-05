/**
 * Text-field placeholder substitution, shared by every binding's text
 * renderer.
 *
 * Pure string logic: resolves OOXML field runs (slide number, date/time,
 * header/footer, document properties, slide title) into their display text.
 * Extracted from the React `viewer/utils/text-field-substitution` module so
 * every binding substitutes identically. The locale-aware datetime engine
 * lives in `text-field-datetime.ts`.
 */

import { formatDatePicture, ooxmlDateTimeFieldPicture } from 'ooxml-core/text';

import { formatLocalizedDateTimeField, isEnglishLocale } from './text-field-datetime';

/** Context for substituting field placeholders (slide number, date/time, header/footer, etc.). */
export interface FieldSubstitutionContext {
	/** 1-based document position of the slide (`PptxSlide.slideNumber`). */
	slideNumber?: number;
	/** `PptxData.firstSlideNumber`; see {@link resolveDisplaySlideNumber}. */
	firstSlideNumber?: number;
	dateTimeText?: string;
	/**
	 * The deck's master date-field default TYPE (`a:fld/@type` on the slide
	 * master's `dt` placeholder, e.g. "datetime2"; see `header-footer-parts.ts`).
	 * This is a field TYPE, not a format pattern, and is used ONLY as a
	 * fallback for a field authored with the generic, unnumbered `datetime`
	 * type - a field that already names its own `datetime1`-`datetime13` type
	 * always uses that type, never this one (see `substituteFieldText`).
	 */
	dateFormat?: string;
	/** Footer text from PptxHeaderFooter settings. */
	footerText?: string;
	/** Header text from PptxHeaderFooter settings. */
	headerText?: string;
	/** Custom document properties for `docproperty` field substitution (keyed by property name). */
	customProperties?: ReadonlyArray<{ name: string; value: string }>;
	/** Locale string for date/time formatting (e.g. "en-US"). Falls back to browser default. */
	locale?: string;
	/** Title text extracted from the first title placeholder on the slide. */
	slideTitle?: string;
}

/**
 * The number a `slidenum` field displays: PowerPoint numbers slide `i`
 * (0-based) as `firstSlideNum + i`, so a deck whose `p:presentation` carries
 * `firstSlideNum="45"` shows 45 on its first slide. `slideNumber` is the
 * 1-based document position; an unset `firstSlideNumber` is the spec default 1.
 * A position below 1 (the master view's pseudo-slide is 0) is not a slide, so
 * it is returned unchanged rather than offset.
 */
export function resolveDisplaySlideNumber(slideNumber: number, firstSlideNumber = 1): number {
	return slideNumber < 1 ? slideNumber : firstSlideNumber + slideNumber - 1;
}

/**
 * Resolve a formatted date string for a given field type.
 *
 * Resolution order:
 * 1. The field's own type, when it names a picture: a predefined
 *    `datetime1`-`datetime13` slug (a non-English `locale` renders it through the
 *    locale's own native convention; English or no locale renders the literal pattern), or a
 *    format-encoded type such as `datetimeyyyy` or `datetimeMMMMyyyy`, whose suffix is the
 *    picture itself (see `ooxmlDateTimeFieldPicture`).
 * 2. `dateFormat` (the deck's master date-field default TYPE) for the generic, unnumbered
 *    `datetime` field type only.
 * 3. The run's own cached text (`cachedText`, what PowerPoint last displayed) for a type that
 *    names no picture, else a locale-aware `toLocaleDateString()`.
 *
 * `dateFormat` is a field TYPE (e.g. "datetime2"), never a raw format pattern: it must not
 * override a field that already names its own type, which is what made every date field on a
 * deck with header/footer date settings render that literal type string.
 */
export function resolveFieldDateText(
	fieldType: string,
	dateFormat?: string,
	locale?: string,
	cachedText?: string,
): string {
	const now = new Date();
	let effectiveType = fieldType;
	if (ooxmlDateTimeFieldPicture(fieldType) === undefined && dateFormat) {
		if (fieldType.toLowerCase() === 'datetime') {
			effectiveType = dateFormat;
		}
	}
	const picture = ooxmlDateTimeFieldPicture(effectiveType);
	if (picture === undefined) {
		// The generic `datetime` type is live by definition, so it never shows stale cached text.
		const generic = fieldType.toLowerCase() === 'datetime';
		return (!generic && cachedText) || now.toLocaleDateString(locale);
	}
	if (!isEnglishLocale(locale)) {
		const localized = formatLocalizedDateTimeField(
			effectiveType.toLowerCase(),
			now,
			locale as string,
		);
		if (localized !== undefined) {
			return localized;
		}
	}
	return formatDatePicture(now, picture, locale === undefined ? {} : { locale });
}

/**
 * Apply field substitution to a text segment if it has a `fieldType`.
 * Returns the substituted text, or the original text if no substitution applies.
 *
 * @param runLocale The field run's own `a:fld/a:rPr@lang`, when the caller has
 *   it (see `TextSegment.style.language`). Takes priority over `ctx.locale`
 *   for every locale-sensitive substitution: a field's own authored language
 *   is more specific than the deck-wide default.
 */
export function substituteFieldText(
	segmentText: string,
	fieldType: string | undefined,
	ctx?: FieldSubstitutionContext,
	runLocale?: string,
): string {
	if (!fieldType || !ctx) {
		return segmentText;
	}
	const fl = fieldType.toLowerCase();
	const locale = runLocale ?? ctx.locale;
	if (fl === 'slidenum' && ctx.slideNumber !== undefined) {
		return String(resolveDisplaySlideNumber(ctx.slideNumber, ctx.firstSlideNumber));
	}
	if (fl.startsWith('datetime')) {
		// Use format-aware date text (the field's own type wins; `dateFormat`
		// only fills in for the generic, unnumbered `datetime` type).
		return resolveFieldDateText(fieldType, ctx.dateFormat, locale, segmentText);
	}
	// Footer field -> resolve from header/footer settings
	if (fl === 'footer' && ctx.footerText !== undefined) {
		return ctx.footerText;
	}
	// Header field -> resolve from header/footer settings
	if (fl === 'header' && ctx.headerText !== undefined) {
		return ctx.headerText;
	}
	// Current date -> system date formatted with locale
	if (fl === 'currentdate') {
		return new Date().toLocaleDateString(locale);
	}
	// Current time -> system time formatted with locale
	if (fl === 'currenttime') {
		return new Date().toLocaleTimeString(locale);
	}
	// Slide title -> resolved from the first title placeholder on the slide
	if (fl === 'slidetitle' && ctx.slideTitle !== undefined) {
		return ctx.slideTitle;
	}
	// Document property -> look up by name from custom properties
	if (fl.startsWith('docproperty') && ctx.customProperties) {
		// Field type format: "docproperty" or "docproperty.PropertyName"
		const dotIdx = fieldType.indexOf('.');
		const propName = dotIdx >= 0 ? fieldType.substring(dotIdx + 1).trim() : '';
		if (propName) {
			const prop = ctx.customProperties.find(
				(p) => p.name.toLowerCase() === propName.toLowerCase(),
			);
			if (prop) {
				return prop.value;
			}
		}
	}
	return segmentText;
}
