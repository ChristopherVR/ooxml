// Locale plumbing shared by the Teams elements. `<teams-app locale="de">` sets its own locale and
// hands it to every Teams element in its shadow tree: an element takes the nearest ancestor's
// locale when it connects, and a host pushes a change down to the elements it rendered. An
// element given its own locale (a standalone `<office-ui-chat-list locale="fr">`) keeps it.
import { teamsI18n, type TeamsI18n } from './translator';

export const LOCALE_HOST = Symbol.for('ooxml-ui.teams.locale-host');

/** An element that has a `locale` and takes part in locale inheritance. */
export interface LocaleHost extends HTMLElement {
	locale: string;
	readonly [LOCALE_HOST]: true;
	readonly renderRoot: HTMLElement | DocumentFragment;
}

const isHost = (node: Node | null): node is LocaleHost =>
	node !== null && (node as unknown as Record<symbol, unknown>)[LOCALE_HOST] === true;

/** The closest ancestor that is a locale host, looking out of shadow roots. */
function ancestorHost(el: Node): LocaleHost | null {
	let node: Node | null = el;
	while (node) {
		const next: Node | null = node.parentNode ?? (node as ShadowRoot).host ?? null;
		if (isHost(next)) return next;
		node = next;
	}
	return null;
}

/** Take the nearest ancestor's locale unless this element was given one. */
export function inheritLocale(el: LocaleHost): void {
	if (el.locale) return;
	const parent = ancestorHost(el);
	if (parent?.locale) el.locale = parent.locale;
}

/** Hand a changed locale to the locale hosts this element rendered. */
export function pushLocale(host: LocaleHost): void {
	for (const child of host.renderRoot.querySelectorAll('*'))
		if (isHost(child) && child.locale !== host.locale) child.locale = host.locale;
}

export const i18nOf = (host: { locale: string }): TeamsI18n => teamsI18n(host.locale);
