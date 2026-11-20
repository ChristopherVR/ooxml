/**
 * compatibility-warning-toasts.ts: turn `handler.getCompatibilityWarnings()`
 * output into the toast list a binding should show.
 *
 * Core's `PptxCompatibilityService` (and a handful of save-path modules that
 * report through it) already flags every markup pattern it silently drops or
 * approximates: unmodelled OOXML it round-trips only via `rawXml`, an
 * external image reference that can't be embedded, signatures stripped on
 * save, a chart's embedded workbook it couldn't write back to, and so on. No
 * binding read `getCompatibilityWarnings()` at all, so an author saving a
 * deck that quietly lost fidelity had no way to know from the UI.
 *
 * This module is deliberately just a lookup: it does not decide WHEN to call
 * `getCompatibilityWarnings()` (after load, after save, ...), only how to
 * turn whatever it returns into toast descriptors with translated copy.
 *
 * Framework-agnostic: no React, Vue, Angular, Svelte or DOM imports.
 */

import type { UpdateRefusalCode } from 'ooxml-core/collab';
import {
	REMOTE_MEDIA_WITHHELD,
	REMOTE_VALUE_REJECTED,
} from 'ooxml-core/pptx/editor/render/collaboration-sync';

/** The `PptxCompatibilityWarning` fields this module reads. Declared structurally. */
export interface CompatibilityWarningSource {
	readonly code: string;
	readonly severity: 'info' | 'warning';
	readonly slideId?: string;
	readonly elementId?: string;
}

export interface CompatibilityWarningToast {
	/** Stable per-code id, since toasts are deduped by code (see module docs). */
	readonly id: string;
	readonly code: string;
	readonly severity: 'info' | 'warning';
	readonly messageKey: string;
	/**
	 * The message already translated, shown instead of looking `messageKey` up (a notice worded
	 * outside the PowerPoint dictionary, such as the signature load notice).
	 */
	readonly text?: string;
	/** `{{code}}` for the generic fallback key; `undefined` for a mapped code. */
	readonly params?: Readonly<Record<string, string>>;
	/**
	 * Messages translated and joined (`pptx.compatibility.partSeparator`) into the `{{parts}}`
	 * param of `messageKey`, for a summary built from counts (the PDF export notice).
	 */
	readonly parts?: readonly ToastMessagePart[];
	/** Translated lines listed under a Details expander. */
	readonly details?: readonly ToastMessagePart[];
}

/** One translatable message of a toast. */
export interface ToastMessagePart {
	readonly key: string;
	readonly params?: Readonly<Record<string, string>>;
}

/**
 * Every `PptxCompatibilityWarning.code` core's compatibility reporter and
 * save-path modules currently emit (grepped across
 * `ooxml-core/src/pptx/core/services/PptxCompatibilityService.ts`,
 * `compatibility-alternate-content.ts`, and every `reportWarning(...)` call
 * site), mapped to its i18n key. A code not in this table falls back to
 * {@link GENERIC_MESSAGE_KEY} with the raw code as a `{{code}}` param, so a
 * future warning code degrades to a readable (if untranslated) toast instead
 * of being silently dropped.
 */
const CODE_MESSAGE_KEYS: Readonly<Record<string, string>> = {
	UNMODELLED_PRESENTATION_MARKUP: 'pptx.compatibility.unmodelledPresentationMarkup',
	UNMODELLED_SLIDE_MARKUP: 'pptx.compatibility.unmodelledSlideMarkup',
	UNMODELLED_SHAPE_PROPERTY: 'pptx.compatibility.unmodelledShapeProperty',
	UNMODELLED_TEXT_BODY_MARKUP: 'pptx.compatibility.unmodelledTextBodyMarkup',
	UNMODELLED_BLIP_FILL_MARKUP: 'pptx.compatibility.unmodelledBlipFillMarkup',
	UNMODELLED_IMAGE_EFFECT: 'pptx.compatibility.unmodelledImageEffect',
	EXTERNAL_IMAGE_REFERENCE: 'pptx.compatibility.externalImageReference',
	SLIDE_SYNCHRONIZATION_METADATA: 'pptx.compatibility.slideSynchronizationMetadata',
	UNSUPPORTED_ALTERNATE_CONTENT_CHOICE: 'pptx.compatibility.unsupportedAlternateContentChoice',
	// The graphic-frame limitations (GRAPHIC_FRAME_LIMITATIONS in core's
	// compatibility-child-sets.ts) had no entry, so their toast showed the
	// generic message with the raw code.
	PARTIAL_OLE_SUPPORT: 'pptx.compatibility.partialOleSupport',
	PARTIAL_INK_SUPPORT: 'pptx.compatibility.partialInkSupport',
	UNSUPPORTED_GRAPHIC_FRAME: 'pptx.compatibility.unsupportedGraphicFrame',
	'group-depth-exceeded': 'pptx.compatibility.groupDepthExceeded',
	SAVE_NOTES_RELATIONSHIP_MISSING: 'pptx.compatibility.saveNotesRelationshipMissing',
	SAVE_NOTES_PART_MISSING: 'pptx.compatibility.saveNotesPartMissing',
	SAVE_NOTES_UPDATE_SKIPPED: 'pptx.compatibility.saveNotesUpdateSkipped',
	CHART_EXTERNAL_DATA_WRITEBACK_UNSUPPORTED:
		'pptx.compatibility.chartExternalDataWritebackUnsupported',
	CHART_EMBEDDED_WORKBOOK_MISSING: 'pptx.compatibility.chartEmbeddedWorkbookMissing',
	CHART_EMBEDDED_WORKBOOK_UNREADABLE: 'pptx.compatibility.chartEmbeddedWorkbookUnreadable',
	CHART_EMBEDDED_WORKBOOK_PARTIAL_WRITEBACK:
		'pptx.compatibility.chartEmbeddedWorkbookPartialWriteback',
	SAVE_IMAGE_PAYLOAD_UNSUPPORTED: 'pptx.compatibility.saveImagePayloadUnsupported',
	SAVE_MEDIA_PAYLOAD_UNSUPPORTED: 'pptx.compatibility.saveMediaPayloadUnsupported',
	SAVE_ELEMENT_SKIPPED: 'pptx.compatibility.saveElementSkipped',
	SAVE_SIGNATURES_STRIPPED: 'pptx.compatibility.saveSignaturesStripped',
	SAVE_GROUP_CHILD_SKIPPED: 'pptx.compatibility.saveGroupChildSkipped',
	SAVE_ANIMATION_SOUND_PAYLOAD_UNSUPPORTED:
		'pptx.compatibility.saveAnimationSoundPayloadUnsupported',
	SAVE_BACKGROUND_IMAGE_UNSUPPORTED: 'pptx.compatibility.saveBackgroundImageUnsupported',
	SHAPE_ID_DEDUPLICATED: 'pptx.compatibility.shapeIdDeduplicated',
	SAVE_TRANSITION_SOUND_PAYLOAD_UNSUPPORTED:
		'pptx.compatibility.saveTransitionSoundPayloadUnsupported',
	DIAGRAM_RELATIONSHIP_IDS_INCOMPLETE: 'pptx.compatibility.diagramRelationshipIdsIncomplete',
	EXPORT_BACKEND_UNAVAILABLE: 'pptx.compatibility.exportBackendUnavailable',
	EMBEDDED_FONT_RESTRICTED: 'pptx.compatibility.embeddedFontRestricted',
	EMBEDDED_FONT_NOT_EMBEDDABLE: 'pptx.compatibility.embeddedFontNotEmbeddable',
	// Recorded as slides are read from a collaboration room, not at load (core's
	// `collaboration-remote-values.ts`); see {@link createRemoteValueNotifier}.
	[REMOTE_VALUE_REJECTED]: 'pptx.compatibility.collaborationValueRejected',
	[REMOTE_MEDIA_WITHHELD]: 'pptx.compatibility.collaborationMediaWithheld',
};

/** Fallback key for a warning code with no entry in {@link CODE_MESSAGE_KEYS}. */
const GENERIC_MESSAGE_KEY = 'pptx.compatibility.generic';

/**
 * Map `getCompatibilityWarnings()` output to the toast list a binding should
 * show, deduped by code.
 *
 * Core already dedupes exact repeats (same code + scope + ids), but a save
 * touching fifty images with the same unsupported payload still reports
 * fifty warnings sharing one code; surfacing fifty identical toasts would be
 * worse than surfacing none, so this collapses to one toast per code. The
 * FIRST warning for a code decides that toast's severity, matching how a
 * user reads a deduped list: the earliest (usually most encompassing) case.
 */
export function compatibilityWarningToasts(
	warnings: readonly CompatibilityWarningSource[],
): CompatibilityWarningToast[] {
	const seen = new Map<string, CompatibilityWarningToast>();
	for (const warning of warnings) {
		if (seen.has(warning.code)) {
			continue;
		}
		const messageKey = CODE_MESSAGE_KEYS[warning.code];
		seen.set(warning.code, {
			id: warning.code,
			code: warning.code,
			severity: warning.severity,
			messageKey: messageKey ?? GENERIC_MESSAGE_KEY,
			params: messageKey ? undefined : { code: warning.code },
		});
	}
	return [...seen.values()];
}

/** A slide's own warnings, structurally. */
interface SlideWarningsSource {
	readonly id?: string;
	readonly warnings?: readonly (CompatibilityWarningSource & { readonly message?: string })[];
}

/**
 * Watch the slides a binding adopts from a collaboration room for values a
 * peer sent that were left out (`COLLABORATION_REMOTE_VALUE_REJECTED`), or for a
 * picture or media file a peer could not share because it is too large
 * (`COLLABORATION_MEDIA_WITHHELD`).
 * Returns a function to call with each newly adopted slide list: it gives the
 * toast to show when the list carries a rejection not reported before, and
 * `undefined` otherwise, so a dismissed toast only returns for new content.
 * Every binding adds the toast to its stack through its usual notice path.
 */
export function createRemoteValueNotifier(): (
	slides: readonly SlideWarningsSource[],
) => CompatibilityWarningToast | undefined {
	const reported = new Set<string>();
	return (slides) => {
		let fresh: CompatibilityWarningSource | undefined;
		for (const warning of remoteValueWarnings(slides)) {
			const key = `${warning.slideId ?? ''}|${warning.elementId ?? ''}|${warning.message ?? ''}`;
			if (!reported.has(key)) {
				reported.add(key);
				fresh ??= warning;
			}
		}
		return fresh ? compatibilityWarningToasts([fresh])[0] : undefined;
	};
}

/** Every collaboration rejection warning `slides` carry. */
export function remoteValueWarnings(
	slides: readonly SlideWarningsSource[],
): (CompatibilityWarningSource & { readonly message?: string })[] {
	return slides.flatMap((slide) =>
		(slide.warnings ?? []).filter(
			(warning) => warning.code === REMOTE_VALUE_REJECTED || warning.code === REMOTE_MEDIA_WITHHELD,
		),
	);
}

/** Toast code of a refused collaboration update; one toast at a time, whatever the reason. */
export const COLLABORATION_REFUSED = 'COLLABORATION_UPDATE_REFUSED';

const REFUSAL_MESSAGE_KEYS: Readonly<Record<UpdateRefusalCode, string>> = {
	'too-large': 'pptx.collaboration.refused.tooLarge',
	malformed: 'pptx.collaboration.refused.malformed',
	'server-rejected': 'pptx.collaboration.refused.serverRejected',
	'server-full': 'pptx.collaboration.refused.serverFull',
};

/**
 * The toast that says why sharing stopped, when an update was refused (the collaboration session
 * controller's `refusal`). Every binding adds it to its notice stack through its usual path.
 */
export function collaborationRefusalToast(code: UpdateRefusalCode): CompatibilityWarningToast {
	return {
		id: COLLABORATION_REFUSED,
		code: COLLABORATION_REFUSED,
		severity: 'warning',
		messageKey: REFUSAL_MESSAGE_KEYS[code],
	};
}
