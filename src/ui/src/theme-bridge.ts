import { OFFICE_TOKENS, type OfficeToken } from './tokens';

/** Legacy shadcn names share one semantic mapping across all product compatibility layers. */
export const LEGACY_OFFICE_THEME_TOKENS = {
	background: '--office-background',
	foreground: '--office-foreground',
	card: '--office-surface',
	'card-foreground': '--office-foreground',
	popover: '--office-popover',
	'popover-foreground': '--office-popover-foreground',
	primary: '--office-accent',
	'primary-foreground': '--office-accent-foreground',
	'primary-ink': '--office-accent-ink',
	secondary: '--office-surface',
	'secondary-foreground': '--office-foreground',
	muted: '--office-surface',
	'muted-foreground': '--office-muted-foreground',
	accent: '--office-selected',
	'accent-foreground': '--office-foreground',
	destructive: '--office-danger',
	'destructive-foreground': '--office-danger-foreground',
	border: '--office-border',
	input: '--office-control-border',
	ring: '--office-ring',
	radius: '--office-radius',
	font: '--office-font',
} as const satisfies Record<string, OfficeToken>;

/**
 * Resolve on an outer host; bridge the computed aliases on inner chrome to avoid CSS cycles. The
 * accent ink follows the product's own primary when a host set only that, so a custom primary
 * keeps colouring its text; otherwise it is the inherited `--office-accent-ink`.
 */
export function legacyThemeAliases(
	selector: string,
	prefix: string,
	privatePrefix: string,
): string {
	return `${selector} {\n${Object.entries(LEGACY_OFFICE_THEME_TOKENS)
		.map(([name, canonical]) => {
			const inherited = `var(${canonical}, ${OFFICE_TOKENS[canonical]})`;
			const value = name === 'primary-ink' ? `var(${prefix}primary, ${inherited})` : inherited;
			return `${privatePrefix}${name}: var(${prefix}${name}, ${value});`;
		})
		.join('\n')}\n}`;
}

export function legacyThemeBridge(selector: string, privatePrefix: string): string {
	return shadcnBridge(selector, privatePrefix, {
		extra: {
			'--office-control-border': `var(${privatePrefix}input)`,
			'--office-font': `var(${privatePrefix}font)`,
			'--office-radius': `var(${privatePrefix}radius)`,
		},
	});
}

/**
 * Feeds the `--office-*` tokens from a product's own theme, so the shared elements follow it.
 * Products (docx, xlsx, pptx, visio) already name their theme after shadcn/ui
 * (`--<prefix>foreground`, `--<prefix>muted-foreground`, ...); `shadcnBridge('--dve-')` is the
 * whole bridge for such a product. `themeBridge` takes any mapping for the rest.
 *
 * The CSS is declared on the selector you give (usually `:host` inside a shadow root, or the
 * product's root element), not on `:root`: a product's colours often resolve where they are read,
 * so a nested theme, a dark scheme or a host override keeps working. Pure strings, no DOM.
 */
export type ThemeBridgeMap = Partial<Record<OfficeToken, string>>;

/** Shadcn names a product may leave out, read from a sibling name instead. */
const SHADCN_FALLBACKS: Readonly<Record<string, string>> = {
	'popover-foreground': 'foreground',
	'primary-ink': 'primary',
};

/** The shadcn token names each `--office-*` colour reads, without the product prefix. */
const SHADCN_NAMES: ReadonlyArray<readonly [OfficeToken, string, string?]> = [
	'foreground',
	'muted-foreground',
	'background',
	'card',
	'accent',
	'border',
	'primary',
	'primary-foreground',
	'primary-ink',
	'ring',
	'popover',
	'popover-foreground',
	'destructive',
	'destructive-foreground',
].map((name) => {
	const token = LEGACY_OFFICE_THEME_TOKENS[name as keyof typeof LEGACY_OFFICE_THEME_TOKENS];
	const fallbackName = SHADCN_FALLBACKS[name];
	return fallbackName ? [token, name, fallbackName] : [token, name];
});

/**
 * `selector { --office-x: value; ... }` for a mapping; entries without a value are skipped. Tokens
 * derived from a mapped one (`--office-select-background` from `--office-background`) need no
 * entry: the theme never declares them, so they resolve where they are read and follow the bridge.
 */
export function themeBridge(selector: string, map: ThemeBridgeMap): string {
	const lines = Object.entries(map)
		.filter((entry): entry is [string, string] => !!entry[1])
		.map(([name, value]) => `\t${name}: ${value};`);
	return `${selector} {\n${lines.join('\n')}\n}\n`;
}

export interface ShadcnBridgeOptions {
	/** Extra or overriding `--office-*` values, for example `{ '--office-font': '...' }`. */
	extra?: ThemeBridgeMap;
	/** A literal each theme variable falls back to when the product does not define it. */
	fallbacks?: Partial<Record<string, string>>;
}

/**
 * The bridge for a product whose theme variables are `<prefix><shadcn-name>`
 * (`prefix` includes the dashes, for example `'--dve-'` or `'--pptx-'`).
 */
export function shadcnBridge(
	selector: string,
	prefix: string,
	options: ShadcnBridgeOptions = {},
): string {
	const map: ThemeBridgeMap = {};
	for (const [token, name, fallbackName] of SHADCN_NAMES) {
		const fallback = options.fallbacks?.[name];
		const inner = fallbackName ? `var(${prefix}${fallbackName})` : fallback;
		map[token] = `var(${prefix}${name}${inner ? `, ${inner}` : ''})`;
	}
	return themeBridge(selector, { ...map, ...options.extra });
}
