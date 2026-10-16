// Shared foundation for the team-workspace elements. Each element is a Lit class with three
// readable parts: reactive properties, a `render()` template and a sibling `.css` file. Properties
// are declared with `declare` and set in the constructor so class fields never shadow Lit's
// accessors. Events are plain bubbling, composed CustomEvents, like every other ooxml-ui control.
import { unsafeCSS, type CSSResult, type PropertyDeclarations, type PropertyValues } from 'lit';
import type { MessageParams } from 'ooxml-core/i18n';
import { OfficeElement, controlStyles } from '../base';
import { inheritLocale, LOCALE_HOST, pushLocale, type LocaleHost } from './i18n/runtime';
import { teamsI18n, type TeamsI18n, type TeamsKey, type TeamsPluralKey } from './i18n/translator';
import resetCss from './reset.css?raw';

/** An element's style list: the shared control base, the team reset and its own .css file. */
export const withStyles = (css: string): CSSResult[] => {
	const [common, own] = controlStyles(css);
	return [common!, unsafeCSS(resetCss), own!];
};

/**
 * Base of the team-workspace elements. They keep Lit's batched (asynchronous) updates: a host
 * awaits `updateComplete`, and a chat list redraws once however many properties it sets.
 *
 * Every element has a `locale` (see `./i18n`): set it on the element, or leave it empty to take
 * the enclosing Teams element's. `t` and `tn` translate in it.
 */
export class TeamsElement extends OfficeElement {
	static override syncUpdates = false;
	static override properties: PropertyDeclarations = { locale: { type: String } };
	/** A BCP 47 tag (`de`, `zh-CN`); empty inherits the enclosing Teams element's, else English. */
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
	/** The accessible name of the host element itself, rewritten with every update. */
	protected hostLabel(): string | undefined {
		return undefined;
	}

	override connectedCallback(): void {
		super.connectedCallback();
		inheritLocale(this as unknown as LocaleHost);
	}

	protected override update(changed: PropertyValues): void {
		super.update(changed);
		const label = this.hostLabel();
		if (label !== undefined) this.setAttribute('aria-label', label);
		if (changed.has('locale')) pushLocale(this as unknown as LocaleHost);
	}
}

/** A stable colour per id, so one person looks the same in every list. */
export function colorFor(id: string): string {
	let hue = 0;
	for (const ch of id) hue = (hue * 31 + ch.charCodeAt(0)) % 360;
	return `hsl(${hue} 45% 42%)`;
}

const DAY_MS = 86_400_000;
export function formatTime(ts: number, locale?: string): string {
	return new Date(ts).toLocaleTimeString(teamsI18n(locale).intl, {
		hour: '2-digit',
		minute: '2-digit',
	});
}
/** Whether a time is today, yesterday or earlier, in the viewer's zone. */
export function dayKind(ts: number, now = Date.now()): 'today' | 'yesterday' | 'other' {
	const start = (t: number): number => new Date(t).setHours(0, 0, 0, 0);
	const diff = Math.round((start(now) - start(ts)) / DAY_MS);
	return diff === 0 ? 'today' : diff === 1 ? 'yesterday' : 'other';
}
/** "Today", "Yesterday" or a date, for the separators between days in a conversation. */
export function dayLabel(ts: number, now = Date.now(), locale?: string): string {
	const i18n = teamsI18n(locale);
	const kind = dayKind(ts, now);
	if (kind === 'today') return i18n.t('time.today');
	if (kind === 'yesterday') return i18n.t('time.yesterday');
	return new Date(ts).toLocaleDateString(i18n.intl, {
		weekday: 'long',
		month: 'long',
		day: 'numeric',
	});
}
export function formatSize(bytes?: number, locale?: string): string {
	if (bytes === undefined) return '';
	const i18n = teamsI18n(locale);
	const number = (n: number, digits = 0): string =>
		new Intl.NumberFormat(i18n.intl, {
			maximumFractionDigits: digits,
			minimumFractionDigits: digits,
		}).format(n);
	if (bytes < 1024) return i18n.t('size.bytes', { size: number(bytes) });
	if (bytes < 1_048_576) return i18n.t('size.kb', { size: number(Math.round(bytes / 1024)) });
	return i18n.t('size.mb', { size: number(bytes / 1_048_576, 1) });
}

const timeOnly = (ts: number, intl: string | undefined): string =>
	new Date(ts).toLocaleTimeString(intl, { hour: 'numeric', minute: '2-digit' });
/** "Thu, Oct 15, 2:00 PM - 3:00 PM" in the viewer's zone (the end shows its date when it differs). */
export function formatWhen(start: number, end: number, locale?: string): string {
	const intl = teamsI18n(locale).intl;
	const day = new Date(start).toLocaleDateString(intl, {
		weekday: 'short',
		month: 'short',
		day: 'numeric',
	});
	const sameDay = new Date(start).toDateString() === new Date(end).toDateString();
	const tail = sameDay
		? timeOnly(end, intl)
		: `${new Date(end).toLocaleDateString(intl, { month: 'short', day: 'numeric' })}, ${timeOnly(end, intl)}`;
	return `${day}, ${timeOnly(start, intl)} - ${tail}`;
}
