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
	secondary: '--office-surface',
	'secondary-foreground': '--office-foreground',
	muted: '--office-surface',
	'muted-foreground': '--office-muted-foreground',
	accent: '--office-selected',
	'accent-foreground': '--office-foreground',
	destructive: '--office-danger',
	'destructive-foreground': '--office-danger-foreground',
	border: '--office-border',
	input: '--office-border',
	ring: '--office-ring',
	radius: '--office-radius',
	font: '--office-font',
} as const satisfies Record<string, OfficeToken>;

/** Resolve on an outer host; bridge the computed aliases on inner chrome to avoid CSS cycles. */
export function legacyThemeAliases(
	selector: string,
	prefix: string,
	privatePrefix: string,
): string {
	return `${selector} {\n${Object.entries(LEGACY_OFFICE_THEME_TOKENS)
		.map(
			([name, canonical]) =>
				`${privatePrefix}${name}: var(${prefix}${name}, var(${canonical}, ${OFFICE_TOKENS[canonical]}));`,
		)
		.join('\n')}\n}`;
}

export function legacyThemeBridge(selector: string, privatePrefix: string): string {
	return shadcnBridge(selector, privatePrefix, {
		extra: {
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
	'ring',
	'popover',
	'popover-foreground',
	'destructive',
	'destructive-foreground',
].map((name) => {
	const token = LEGACY_OFFICE_THEME_TOKENS[name as keyof typeof LEGACY_OFFICE_THEME_TOKENS];
	return name === 'popover-foreground' ? [token, name, 'foreground'] : [token, name];
});

/** The `--office-*` tokens a token's default reads, such as `--office-background` for selects. */
const reads = (value: string) => [...value.matchAll(/var\((--office-[\w-]+)/g)].map((m) => m[1]!);

/**
 * The tokens whose defaults read (directly or through another default) one of `bridged`. The
 * theme declares those defaults on `:root`, and a `var()` resolves where it is declared, so a
 * bridge that only re-points `--office-background` would leave `--office-select-background`
 * holding the page's colour (black selects on a light viewer when the system is dark).
 */
function derivedTokens(bridged: ReadonlySet<string>): [string, string][] {
	const derived = new Map<string, string>();
	let grew = true;
	while (grew) {
		grew = false;
		for (const [name, value] of Object.entries(OFFICE_TOKENS)) {
			if (bridged.has(name) || derived.has(name)) continue;
			if (reads(value).some((token) => bridged.has(token) || derived.has(token))) {
				derived.set(name, value);
				grew = true;
			}
		}
	}
	return [...derived];
}

/**
 * `selector { --office-x: value; ... }` for a mapping; entries without a value are skipped. The
 * tokens derived from a mapped one are declared again on the selector with their defaults, so they
 * follow the product's colours instead of the page's.
 */
export function themeBridge(selector: string, map: ThemeBridgeMap): string {
	const mapped = Object.entries(map).filter((entry): entry is [string, string] => !!entry[1]);
	const lines = [...mapped, ...derivedTokens(new Set(mapped.map(([name]) => name)))].map(
		([name, value]) => `\t${name}: ${value};`,
	);
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
