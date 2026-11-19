import { externalContentBarVisible, resolveAllowExternalContent } from 'ooxml-ui/pptx';
import type { ViewerOptions } from 'ooxml-ui/pptx';
/**
 * useExternalContentConsent: PowerPoint's security bar for linked pictures and
 * media (shared `external-content-policy`), kept per presentation.
 *
 * Consent belongs to the `content` it was given for, so opening another
 * presentation (host prop, File > Open, a version or recovery copy) starts
 * blocked again without a reset effect, while the reload Enable Content asks
 * for keeps the same `content` and its consent.
 */
import { useCallback, useState } from 'react';

type Content = ArrayBuffer | Uint8Array;

/** What a finished load reports: whether the deck links web content, and whether it was fetched. */
export interface LinkedWebContentLoad {
	hasLinkedWebContent: boolean;
	allowed: boolean;
}

/**
 * Enable Content's reload of the presentation `content`: from `bytes` (the
 * source, or the serialized deck when it had unsaved changes), keeping it
 * marked unsaved when it was, on the slide that was showing.
 */
export interface ExternalContentReload {
	content: Content;
	bytes: Content;
	keepDirty: boolean;
	slideIndex: number;
}

/** What Enable Content reads from the viewer when it runs. */
export interface ExternalContentEnableSource {
	isDirty: boolean;
	activeSlideIndex: number;
	/** The deck as the viewer serializes it for recovery (plaintext, with unsaved edits). */
	serialize: () => Promise<Uint8Array | null>;
}

interface Consent {
	content: Content | null;
	enabled: boolean;
	closed: boolean;
}

const NO_CONSENT: Consent = { content: null, enabled: false, closed: false };

export function useExternalContentConsent(
	content: Content | null,
	host: boolean | undefined,
	options: Pick<ViewerOptions, 'trust'>,
) {
	const [consent, setConsent] = useState<Consent>(NO_CONSENT);
	const [load, setLoad] = useState<(LinkedWebContentLoad & { content: Content | null }) | null>(
		null,
	);
	const [reload, setReload] = useState<ExternalContentReload | null>(null);

	const own = content !== null && consent.content === content;
	const enabled = own && consent.enabled;
	const closed = own && consent.closed;
	const current = load && load.content === content ? load : null;

	const reportLoad = useCallback(
		(value: LinkedWebContentLoad) => setLoad({ ...value, content }),
		[content],
	);
	const enable = useCallback(
		(source: ExternalContentEnableSource) => {
			if (!content) {
				return;
			}
			setConsent({ content, enabled: true, closed: false });
			void (async () => {
				const serialized = source.isDirty ? await source.serialize() : null;
				setReload({
					content,
					bytes: serialized ?? content,
					keepDirty: source.isDirty,
					slideIndex: source.activeSlideIndex,
				});
			})();
		},
		[content],
	);
	const close = useCallback(() => {
		if (content) {
			setConsent((previous) => ({
				content,
				enabled: previous.content === content && previous.enabled,
				closed: true,
			}));
		}
	}, [content]);

	return {
		/** The main load's `allowExternalImages`: host, Trust Center, or Enable Content here. */
		allowExternalImages: resolveAllowExternalContent(host, options, enabled),
		/** Whether the security bar shows for the loaded presentation. */
		barVisible: externalContentBarVisible({
			hasLinkedWebContent: current?.hasLinkedWebContent,
			allowed: (current?.allowed ?? false) || enabled,
			host,
			closed,
		}),
		/** The pending Enable Content reload, for `useLoadContent`. */
		reload,
		reportLoad,
		enable,
		close,
	};
}

export type ExternalContentConsent = ReturnType<typeof useExternalContentConsent>;
