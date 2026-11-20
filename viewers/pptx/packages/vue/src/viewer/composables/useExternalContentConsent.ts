/**
 * useExternalContentConsent: PowerPoint's security bar for linked pictures and
 * media (shared `external-content-policy`), kept per presentation and per
 * viewer.
 *
 * Consent belongs to the content source it was given for, so opening another
 * presentation (host `content`, File > Open, a version or recovery copy)
 * starts blocked again without a reset racing the load. Enable Content only
 * changes what this viewer draws (`provideLinkedContent`): the deck, its
 * unsaved state, undo history, autosave and a collaboration session are
 * untouched, and each collaborator decides for themselves.
 */
import { resolveExternalContentView } from 'ooxml-ui/pptx';
import type { ExternalContentOptions, ViewerExternalContentBar } from 'ooxml-ui/pptx';
import type { ComputedRef } from 'vue';
import { computed, shallowRef } from 'vue';

export interface UseExternalContentConsentInput {
	/** The presentation source the consent is for. */
	content: () => unknown;
	/** The host `allowExternalContent` switch. */
	host: () => boolean | undefined;
	/** The viewer's options (Trust Center > "Allow external content"). */
	options: () => ExternalContentOptions;
	/**
	 * What the viewer draws (its slides, template elements, table styles), read for linked
	 * content: a deck adopted from a collaboration room counts as much as a loaded one.
	 */
	drawn: () => unknown;
}

interface Consent {
	content: unknown;
	enabled: boolean;
	closed: boolean;
}

export function useExternalContentConsent(
	input: UseExternalContentConsentInput,
): ComputedRef<ViewerExternalContentBar> {
	const consent = shallowRef<Consent>({ content: undefined, enabled: false, closed: false });
	const own = (): Consent => {
		const content = input.content();
		return consent.value.content === content
			? consent.value
			: { content, enabled: false, closed: false };
	};
	const enable = (): void => {
		consent.value = { ...own(), enabled: true };
	};
	const close = (): void => {
		consent.value = { ...own(), closed: true };
	};
	return computed(() => {
		const current = own();
		const view = resolveExternalContentView({
			host: input.host(),
			options: input.options(),
			enabled: current.enabled,
			closed: current.closed,
			content: input.drawn(),
		});
		return { ...view, enable, close };
	});
}
