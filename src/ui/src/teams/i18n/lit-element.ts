// Base of the Teams app elements that are plain Lit elements (the app, its dialogs and panes).
// `TeamsElement` (../base) is the same contract over `OfficeElement`.
import { LitElement, type PropertyDeclarations, type PropertyValues } from 'lit';
import type { MessageParams } from 'ooxml-core/i18n';
import { inheritLocale, LOCALE_HOST, pushLocale, type LocaleHost } from './runtime';
import { teamsI18n, type TeamsI18n, type TeamsKey, type TeamsPluralKey } from './translator';

export class TeamsLitElement extends LitElement {
	static override properties: PropertyDeclarations = { locale: { type: String } };
	/** A BCP 47 tag; empty inherits the enclosing Teams element's, else English. */
	declare locale: string;

	constructor() {
		super();
		this.locale = '';
	}

	get [LOCALE_HOST](): true {
		return true;
	}

	protected get i18n(): TeamsI18n {
		return teamsI18n(this.locale);
	}
	protected t(key: TeamsKey, params?: MessageParams): string {
		return this.i18n.t(key, params);
	}
	protected tn(key: TeamsPluralKey, count: number, params?: MessageParams): string {
		return this.i18n.tn(key, count, params);
	}

	override connectedCallback(): void {
		super.connectedCallback();
		inheritLocale(this as unknown as LocaleHost);
	}

	protected override update(changed: PropertyValues): void {
		super.update(changed);
		if (changed.has('locale')) pushLocale(this as unknown as LocaleHost);
	}
}
