/**
 * Seed and style the runs of a DOM inline text editor (vanilla, Svelte) so the
 * text keeps the face, size and spacing view mode painted it in.
 *
 * Those editors used to seed bare text (or unstyled spans) and set the body's
 * RAW `fontFamily` on the surface, overriding the block style. The raw name has
 * no substitution chain, so when the deck font was not painted (a webfont still
 * downloading, offline, or the download option off) view mode fell back to the
 * chain's stand-in while the editor fell back to the browser default serif:
 * the font visibly changed on double-click. Each run now carries the same
 * {@link segmentStyleToCss} style view mode and the list editor use, inside the
 * body's inherited font (substituted, autofit-scaled).
 *
 * @module render/inline-text-seed
 */
import type { TextSegment, TextStyle } from 'ooxml-core/pptx';

import { isBulletMarkerSegment } from './bullet-toggle';
import { applyUnderlineVariant } from './text-run-decoration';
import { segmentStyleToCss } from './text-run-style';
import { resolveBlockFont } from './text-paragraphs';
import { resolveAutoFitFontScale } from './text-style-helpers';

/** Write `segment`'s view-mode run style onto `span` (body font inherited, autofit applied). */
export function applyInlineRunStyle(
	span: HTMLElement,
	segment: TextSegment,
	textStyle: TextStyle | undefined,
): void {
	const fontScale = resolveAutoFitFontScale(textStyle);
	const style = segmentStyleToCss(segment, fontScale, {
		blockFont: resolveBlockFont(textStyle, fontScale),
	});
	applyUnderlineVariant(style, segment);
	for (const [key, value] of Object.entries(style)) {
		if (value === undefined || value === null || value === '') {
			continue;
		}
		span.style.setProperty(
			key.replace(/[A-Z]/g, (char) => `-${char.toLowerCase()}`),
			typeof value === 'number' && !UNITLESS.has(key) ? `${value}px` : String(value),
		);
	}
}

/** Run style keys whose numbers are not lengths. */
const UNITLESS = new Set(['fontWeight', 'lineHeight', 'opacity', 'zIndex', 'flexGrow']);

/**
 * Seed a plain (non-list) inline editor with `segments`, one styled span per
 * run inside a single text-flow container, or with `initialText` when the
 * element has no segments. Returns the element that holds the text.
 */
export function seedPlainInlineText(
	surface: HTMLElement,
	segments: readonly TextSegment[] | undefined,
	initialText: string,
	textStyle?: TextStyle,
): HTMLElement {
	const doc = surface.ownerDocument;
	if (!segments?.length) {
		surface.textContent = initialText;
		return surface;
	}
	// The text-block style makes the surface a flex column so vertical
	// alignment applies to paragraphs. Keep rich-text runs inside one flex
	// item; direct flex children are blockified onto separate rows.
	const textFlow = doc.createElement('div');
	textFlow.dataset.pptxTextFlow = '';
	segments.forEach((segment, index) => {
		const span = doc.createElement('span');
		span.dataset.segIdx = String(index);
		applyInlineRunStyle(span, segment, textStyle);
		if (isBulletMarkerSegment(segment)) {
			span.dataset.pptxBulletMarker = '';
			span.contentEditable = 'false';
		}
		const precedingMarker = index > 0 && isBulletMarkerSegment(segments[index - 1]);
		const carriesList = segment.bulletInfo && !segment.bulletInfo.none;
		if (
			segment.text.length === 0 &&
			index === segments.length - 1 &&
			(precedingMarker || carriesList)
		) {
			// An empty inline span after a non-editable list marker has no caret
			// position. A display-only BR lets that pending list item receive text.
			span.dataset.pptxEmptyRun = '';
			span.appendChild(doc.createElement('br'));
		} else {
			span.textContent = segment.text;
		}
		textFlow.appendChild(span);
	});
	surface.appendChild(textFlow);
	return textFlow;
}
