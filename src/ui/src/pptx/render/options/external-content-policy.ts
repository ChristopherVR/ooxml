/**
 * Linked external content: whether a deck may make the browser contact other
 * servers for pictures and media it links (`TargetMode="External"`, http(s))
 * instead of embedding.
 *
 * The default is Microsoft PowerPoint's (see `docs/powerpoint-external-content.md`):
 * nothing is downloaded when a deck opens, and a security bar offers
 * "Enable Content" for that presentation. Enabling lasts until another
 * presentation is opened and is never stored.
 *
 * The host `allowExternalContent` switch overrides it: `true` downloads without
 * asking (PowerPoint's "Unblock automatic download of linked images" policy),
 * `false` blocks without the bar. Either locks the Trust Center option
 * `trust.allowExternalContent`, which is off by default; a user who turns it
 * on downloads without asking. Every binding feeds the result of
 * {@link resolveAllowExternalContent} to the loader's `allowExternalImages`
 * option, shows the bar when {@link externalContentBarVisible} says so, reloads
 * from {@link prepareExternalContentReload} on Enable Content and folds
 * {@link lockExternalContentSetting} into the options locks; no rule is
 * re-derived per binding.
 *
 * @module render/options/external-content-policy
 */
import type { OptionsSettingId, OptionsSettingValues } from '../customization/customization-types';
import type { ViewerOptions } from './viewer-options';

/** The `<group>.<key>` id of the Trust Center setting the host switch locks. */
export const EXTERNAL_CONTENT_SETTING_ID: OptionsSettingId = 'trust.allowExternalContent';

/**
 * Whether a load may fetch linked external content: the host value when set,
 * otherwise the user's Trust Center option or Enable Content on the security
 * bar for this presentation (`enabledForDocument`).
 */
export function resolveAllowExternalContent(
	host: boolean | undefined,
	options: Pick<ViewerOptions, 'trust'>,
	enabledForDocument = false,
): boolean {
	if (typeof host === 'boolean') {
		return host;
	}
	return options.trust.allowExternalContent || enabledForDocument;
}

/** What decides whether the security bar shows for the loaded deck. */
export interface ExternalContentBarInput {
	/** The load result's `hasLinkedWebContent`. */
	hasLinkedWebContent: boolean | undefined;
	/** What the load used, from {@link resolveAllowExternalContent}. */
	allowed: boolean;
	/** The host `allowExternalContent` switch. */
	host: boolean | undefined;
	/** The user closed the bar for this presentation. */
	closed: boolean;
}

/**
 * PowerPoint's security bar: the deck links web content this load left out,
 * the host has not blocked it outright, and the user has not closed the bar.
 */
export function externalContentBarVisible(input: ExternalContentBarInput): boolean {
	return (
		input.hasLinkedWebContent === true && !input.allowed && input.host !== false && !input.closed
	);
}

/** What Enable Content reads from the viewer at the moment the user chooses it. */
export interface ExternalContentEnableSource {
	/**
	 * The presentation as it stands, saved and unsaved edits included, in bytes the
	 * viewer can read back (the recovery serialisation: plaintext, never encrypted).
	 * Null when no deck is loaded.
	 */
	serializeCurrentDeck(): Promise<Uint8Array | null>;
	/** The deck has unsaved changes; they stay unsaved after the reload. */
	isDirty: boolean;
	/** The slide on screen, which the reload keeps. */
	slideIndex: number;
}

/** The reload Enable Content asks a binding to make, with linked content allowed. */
export interface ExternalContentReload {
	/** The current deck, from {@link ExternalContentEnableSource.serializeCurrentDeck}. */
	bytes: Uint8Array;
	/** Mark the reloaded deck unsaved, as it was. */
	keepDirty: boolean;
	/** The slide to show after the reload, clamped by the binding. */
	slideIndex: number;
}

/**
 * Enable Content's reload. It is always the deck as it now stands: the bytes it
 * was opened from would bring back what was edited since, saved or not, and the
 * next save would write that old deck over the saved one. Null when there is
 * no deck to serialise, in which case nothing reloads.
 *
 * The reload starts undo history over: the steps before it hold the deck
 * without its linked content, so undoing into them would remove it again.
 */
export async function prepareExternalContentReload(
	source: ExternalContentEnableSource,
): Promise<ExternalContentReload | null> {
	const bytes = await source.serializeCurrentDeck();
	return bytes ? { bytes, keepDirty: source.isDirty, slideIndex: source.slideIndex } : null;
}

/**
 * The security bar as every binding's building blocks hand it to a custom shell,
 * which renders the binding's `ExternalContentBar` while `barVisible` and wires
 * its Enable Content and Close to `enable` and `close`.
 */
export interface ViewerExternalContentBar {
	/** The deck links pictures or media to the web, and they were not downloaded. */
	barVisible: boolean;
	/** Enable Content: reload this presentation as it stands, with its linked content. */
	enable: () => void;
	/** Close the bar; the content stays blocked. */
	close: () => void;
}

/** The bar's two requests, as its `external-content-request` event carries them. */
export type ExternalContentBarIntent = { id: 'action' | 'close' };

/** A host's handler per bar request; a request without one does nothing. */
export interface ExternalContentBarIntentHandlers {
	/** Enable Content: reload this presentation with its linked content. */
	enable?: (() => void) | undefined;
	/** Close: hide the bar and keep the content blocked. */
	close?: (() => void) | undefined;
}

/** Route the bar's `external-content-request` event to the host's handler for it. */
export function dispatchExternalContentBarIntent(
	event: Event,
	handlers: ExternalContentBarIntentHandlers,
): void {
	const intent = (event as CustomEvent<ExternalContentBarIntent | undefined>).detail;
	if (intent?.id === 'action') {
		handlers.enable?.();
	} else if (intent?.id === 'close') {
		handlers.close?.();
	}
}

/**
 * The options locks with the host switch folded in. Returns `locked` itself
 * when the host left the switch unset, and a new record otherwise; the host
 * switch wins over a `customization.options.locked` entry for the same id.
 */
export function lockExternalContentSetting(
	locked: OptionsSettingValues,
	host: boolean | undefined,
): OptionsSettingValues {
	if (typeof host !== 'boolean') {
		return locked;
	}
	return { ...locked, [EXTERNAL_CONTENT_SETTING_ID]: host };
}
