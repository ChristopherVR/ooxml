/**
 * The host `limits` option of every viewer binding: the message a binding's
 * existing load-error UI shows when a deck breaks one of the core's
 * `PptxLoadLimits`.
 *
 * Each binding passes its `limits` value straight to `PptxHandler.load` (unset
 * keeps the core defaults) and calls {@link pptxLoadErrorMessage} from its load
 * `catch`, so the wording and the number formatting live here once.
 *
 * @module render/load-limits
 */
import { getPptxLoadLimitViolation } from 'ooxml-core/pptx';
import type { PptxLoadLimitName, PptxLoadLimits } from 'ooxml-core/pptx';

export type { PptxLoadLimits };

/** The `t(key, params)` call {@link pptxLoadErrorMessage} makes; every binding's translator fits it. */
export type LoadLimitTranslate = (
	key: string,
	params: { readonly actual: string; readonly maximum: string },
) => string;

/** The i18n key of the message for each limit. */
export const LOAD_LIMIT_MESSAGE_KEYS: Readonly<Record<PptxLoadLimitName, string>> = Object.freeze({
	maxInputBytes: 'pptx.viewer.loadLimit.maxInputBytes',
	maxUncompressedBytes: 'pptx.viewer.loadLimit.maxUncompressedBytes',
	maxEntries: 'pptx.viewer.loadLimit.maxEntries',
	maxSlides: 'pptx.viewer.loadLimit.maxSlides',
	maxShapes: 'pptx.viewer.loadLimit.maxShapes',
});

const BYTE_LIMITS: ReadonlySet<PptxLoadLimitName> = new Set([
	'maxInputBytes',
	'maxUncompressedBytes',
]);

/** A byte count as `512 B`, `12.5 KB` or `3.2 MB` (binary units, at most one decimal). */
export function formatLoadLimitBytes(bytes: number): string {
	if (bytes < 1024) {
		return `${Math.round(bytes)} B`;
	}
	const kib = bytes / 1024;
	if (kib < 1024) {
		return `${Math.round(kib * 10) / 10} KB`;
	}
	return `${Math.round((kib / 1024) * 10) / 10} MB`;
}

function formatValue(limit: PptxLoadLimitName, value: number): string {
	return BYTE_LIMITS.has(limit) ? formatLoadLimitBytes(value) : String(Math.round(value));
}

/**
 * The text a binding shows for a failed load: for a load-limit violation, the
 * localized `pptx.viewer.loadLimit.<limit>` message with the observed value
 * (`{{actual}}`) and the limit (`{{maximum}}`); otherwise the error's own message.
 */
export function pptxLoadErrorMessage(error: unknown, t: LoadLimitTranslate): string {
	const violation = getPptxLoadLimitViolation(error);
	if (violation) {
		const { limit, maximum } = violation;
		// Without an observed value, the deck is known to be over the limit only.
		const actual = violation.actual ?? maximum + 1;
		return t(LOAD_LIMIT_MESSAGE_KEYS[limit], {
			actual: formatValue(limit, actual),
			maximum: formatValue(limit, maximum),
		});
	}
	return error instanceof Error ? error.message : String(error);
}
