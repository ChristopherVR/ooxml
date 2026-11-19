/**
 * Linked external content: whether a viewer may make the browser contact other
 * servers for pictures and media a deck links (`TargetMode="External"`, http(s))
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
 * on downloads without asking.
 *
 * The deck's model always keeps the linked addresses (viewers load with
 * `allowExternalImages: true`); what is fetched is decided per viewer, where the
 * model is drawn (`linked-content`). Every binding derives `allowed` and the bar
 * from {@link resolveExternalContentView}, draws through `linkedContentView` and
 * folds {@link lockExternalContentSetting} into the options locks; no rule is
 * re-derived per binding.
 *
 * @module render/options/external-content-policy
 */
import type { OptionsSettingId, OptionsSettingValues } from '../customization/customization-types';
import { hasLinkedContent } from './linked-content';
import type { ViewerTrustOptions } from './viewer-options';

/** The one option the rule reads: Trust Center > "Allow external content". */
export type ExternalContentOptions = { trust: Pick<ViewerTrustOptions, 'allowExternalContent'> };

/** The `<group>.<key>` id of the Trust Center setting the host switch locks. */
export const EXTERNAL_CONTENT_SETTING_ID: OptionsSettingId = 'trust.allowExternalContent';

/**
 * Whether a viewer may fetch linked external content: the host value when set,
 * otherwise the user's Trust Center option or Enable Content on the security
 * bar for this presentation (`enabledForDocument`).
 */
export function resolveAllowExternalContent(
	host: boolean | undefined,
	options: ExternalContentOptions,
	enabledForDocument = false,
): boolean {
	if (typeof host === 'boolean') {
		return host;
	}
	return options.trust.allowExternalContent || enabledForDocument;
}

/** What decides what a viewer shows of the deck's linked web content. */
export interface ExternalContentViewInput {
	/** The host `allowExternalContent` switch. */
	host: boolean | undefined;
	/** The viewer's options; `trust.allowExternalContent` is the Trust Center switch. */
	options: ExternalContentOptions;
	/** The user chose Enable Content for this presentation. */
	enabled: boolean;
	/** The user closed the bar for this presentation. */
	closed: boolean;
	/**
	 * What the viewer draws: its slides, and anything else of the deck it renders (template
	 * elements, the table style map). Read with `hasLinkedContent`.
	 */
	content: unknown;
}

/** What a viewer shows of the deck's linked web content. */
export interface ExternalContentView {
	/** Surfaces draw the content itself (`linkedContentView(value, allowed)`). */
	allowed: boolean;
	/**
	 * PowerPoint's security bar: the deck links web content this viewer withholds, the host
	 * has not blocked it outright and the user has not closed the bar.
	 */
	barVisible: boolean;
}

/**
 * The one rule every binding and building block follows. Consent is this viewer's own: it is
 * never stored in the deck or sent to collaborators, so each participant decides for itself,
 * and Enable Content only redraws.
 */
export function resolveExternalContentView(input: ExternalContentViewInput): ExternalContentView {
	const allowed = resolveAllowExternalContent(input.host, input.options, input.enabled);
	return {
		allowed,
		barVisible:
			!allowed && input.host !== false && !input.closed && hasLinkedContent(input.content),
	};
}

/**
 * The security bar as every binding's building blocks hand it to a custom shell,
 * which renders the binding's `ExternalContentBar` while `barVisible` and wires
 * its Enable Content and Close to `enable` and `close`.
 */
export interface ViewerExternalContentBar {
	/** The deck links pictures or media to the web, and they were not downloaded. */
	barVisible: boolean;
	/**
	 * This viewer draws the linked content. A shell that draws slides itself (its own
	 * thumbnails) reads them through `linkedContentView(value, allowed)`.
	 */
	allowed: boolean;
	/** Enable Content: draw this presentation's linked content, for this viewer only. */
	enable: () => void;
	/** Close the bar; the content stays blocked. */
	close: () => void;
}

/** The bar's two requests, as its `external-content-request` event carries them. */
export type ExternalContentBarIntent = { id: 'action' | 'close' };

/** A host's handler per bar request; a request without one does nothing. */
export interface ExternalContentBarIntentHandlers {
	/** Enable Content: draw this presentation's linked content. */
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
