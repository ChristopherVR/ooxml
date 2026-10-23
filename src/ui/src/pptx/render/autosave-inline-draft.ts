/**
 * In-progress inline text for recovery autosave.
 *
 * While a text box is being typed into, the typed text lives only in the
 * inline editor: it reaches the slides when the edit is committed (blur,
 * Escape-commit, click away). PowerPoint's AutoRecover still captures typing
 * that was never committed, so a crash mid-sentence is recoverable. The five
 * bindings already report every keystroke through one place, the live
 * collaboration preview; {@link reportInlineTextDraft} is that one place for
 * autosave too: it publishes the preview and hands the draft to the autosave
 * engine (`AutosaveController.noteInlineDraft`), which counts it as an edit and
 * passes it to the snapshot serialiser, where {@link applyInlineTextDraft}
 * writes it into the slides being saved.
 *
 * @module render/autosave-inline-draft
 */
import { hasTextProperties } from 'ooxml-core/pptx';
import type { PptxSlide } from 'ooxml-core/pptx';

import type { CollaborationLivePatcher } from './collaboration-live-patch';
import { publishLiveInlineText } from './collaboration-live-patch';
import { remapTextToSegments } from './remap-text';

/** The text an inline editor holds that the slides do not have yet. */
export interface InlineTextDraft {
	/** The slide the edited element is on, when known. */
	readonly slideId?: string | undefined;
	readonly elementId: string;
	readonly text: string;
}

/** What receives drafts: the shared autosave engine. */
export interface InlineTextDraftSink {
	noteInlineDraft(draft: InlineTextDraft | null): void;
}

/**
 * Report one keystroke of inline text: the collaboration preview and the
 * autosave draft. `elementId` null (the editor closed) clears the draft.
 */
export function reportInlineTextDraft(
	targets: {
		readonly livePatcher?: CollaborationLivePatcher | null | undefined;
		readonly autosave?: InlineTextDraftSink | null | undefined;
	},
	slide: PptxSlide | undefined,
	elementId: string | null | undefined,
	text: string,
): void {
	publishLiveInlineText(targets.livePatcher, slide, elementId, text);
	targets.autosave?.noteInlineDraft(elementId ? { slideId: slide?.id, elementId, text } : null);
}

/**
 * The slides with `draft` written into its element (text and remapped runs),
 * or the same array when there is nothing to apply.
 */
export function applyInlineTextDraft<S extends PptxSlide>(
	slides: readonly S[],
	draft: InlineTextDraft | null | undefined,
): readonly S[] {
	if (!draft) {
		return slides;
	}
	let changed = false;
	const next = slides.map((slide) => {
		if (changed || (draft.slideId !== undefined && slide.id !== draft.slideId)) {
			return slide;
		}
		const index = slide.elements.findIndex((element) => element.id === draft.elementId);
		const element = index >= 0 ? slide.elements[index] : undefined;
		if (!element || !hasTextProperties(element) || element.text === draft.text) {
			return slide;
		}
		changed = true;
		const elements = [...slide.elements];
		elements[index] = {
			...element,
			text: draft.text,
			textSegments: remapTextToSegments(draft.text, element.textSegments, element.textStyle),
		};
		return { ...slide, elements };
	});
	return changed ? next : slides;
}
