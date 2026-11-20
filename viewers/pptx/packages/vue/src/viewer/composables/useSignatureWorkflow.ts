import { signatureStatusView } from 'ooxml-ui/pptx';
import type { SignatureStatusView } from 'ooxml-ui/pptx';
import type { SignatureSummary } from 'pptx-viewer-core';
import { computed, ref, toValue, watch } from 'vue';
import type { ComputedRef, MaybeRefOrGetter, Ref } from 'vue';

export interface UseSignatureWorkflowInput {
	/** The loaded deck's signatures (`PptxData.signatures`), absent when it carries none. */
	signatures: Ref<SignatureSummary | undefined>;
	/** The viewer's UI locale (vue-i18n `locale`): the view's wording follows it. */
	locale: MaybeRefOrGetter<string | undefined>;
	/** Autosave dirty flag; a deck carrying signatures warns once on its first dirtying edit. */
	isDirty: ComputedRef<boolean> | Ref<boolean>;
}

export interface UseSignatureWorkflowResult {
	showSignatures: Ref<boolean>;
	/** How every signature surface presents the deck: signed only when its signatures verify. */
	view: ComputedRef<SignatureStatusView>;
	showSignatureStripped: Ref<boolean>;
	onAckSignatureStripped: () => void;
}

/**
 * useSignatureWorkflow: File > Digital Signatures panel plus the "saving a
 * signed deck strips signatures" first-edit warning (mirrors React's
 * `useViewerDialogs` signature-strip effect). The badge, panel and warning all
 * read the one shared `signatureStatusView`.
 */
export function useSignatureWorkflow(input: UseSignatureWorkflowInput): UseSignatureWorkflowResult {
	const { signatures, isDirty } = input;

	const showSignatures = ref(false);
	const view = computed(() => signatureStatusView(signatures.value, toValue(input.locale)));
	// Warn once, on the first edit of a deck carrying signatures, that saving strips them.
	const showSignatureStripped = ref(false);
	const signatureStripAcknowledged = ref(false);
	watch(
		() => isDirty.value,
		(dirty) => {
			if (dirty && view.value.present && !signatureStripAcknowledged.value) {
				showSignatureStripped.value = true;
			}
		},
	);
	function onAckSignatureStripped(): void {
		signatureStripAcknowledged.value = true;
		showSignatureStripped.value = false;
	}

	return { showSignatures, view, showSignatureStripped, onAckSignatureStripped };
}
