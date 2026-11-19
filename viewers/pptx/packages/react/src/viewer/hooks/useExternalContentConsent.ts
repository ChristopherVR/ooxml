import { resolveExternalContentView } from 'ooxml-ui/pptx';
import type { ViewerOptions } from 'ooxml-ui/pptx';
/**
 * useExternalContentConsent: PowerPoint's security bar for linked pictures and
 * media (shared `external-content-policy`), kept per presentation and per
 * viewer.
 *
 * Consent belongs to the `content` it was given for, so opening another
 * presentation (host prop, File > Open, a version or recovery copy) starts
 * blocked again without a reset effect. Enable Content only changes what this
 * viewer draws (`LinkedContentContext`): the deck, its unsaved state, undo
 * history, autosave and a collaboration session are untouched, and a
 * collaborator decides for themselves.
 */
import { useCallback, useState } from 'react';

type Content = ArrayBuffer | Uint8Array | null;

interface Consent {
	content: Content;
	enabled: boolean;
	closed: boolean;
}

const NO_CONSENT: Consent = { content: null, enabled: false, closed: false };

/**
 * @param content The presentation source the consent is for.
 * @param drawn What the viewer draws (its slides, template elements, table styles), read for
 *   linked content: a deck adopted from a collaboration room counts as much as a loaded one.
 */
export function useExternalContentConsent(
	content: Content,
	host: boolean | undefined,
	options: Pick<ViewerOptions, 'trust'>,
	drawn: unknown,
) {
	const [consent, setConsent] = useState<Consent>(NO_CONSENT);
	const own = consent.content === content;
	const view = resolveExternalContentView({
		host,
		options,
		enabled: own && consent.enabled,
		closed: own && consent.closed,
		content: drawn,
	});
	const enable = useCallback(
		() => setConsent((previous) => ({ ...ownConsent(previous, content), enabled: true })),
		[content],
	);
	const close = useCallback(
		() => setConsent((previous) => ({ ...ownConsent(previous, content), closed: true })),
		[content],
	);
	return { ...view, enable, close };
}

function ownConsent(previous: Consent, content: Content): Consent {
	return previous.content === content ? previous : { ...NO_CONSENT, content };
}

export type ExternalContentConsent = ReturnType<typeof useExternalContentConsent>;
