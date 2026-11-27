// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { ParagraphStyleCatalog, TextRun } from './model';
import type { RunFormatting, RunStyleCatalog } from './run-style-model';
import { cachedResolver, chainResolver, defaultStyleId, mergeDefined } from './style-inheritance';

/**
 * Word "toggle" run properties (ECMA-376 17.7.3): each level in the formatting
 * hierarchy that specifies the property flips its effective state, regardless
 * of the level's own boolean value, so two ancestor levels that both turn
 * bold "on" cancel back out to "off" (Word's documented double-toggle quirk).
 */
const TOGGLE_KEYS = [
	'bold',
	'italic',
	'boldComplexScript',
	'italicComplexScript',
	'caps',
	'smallCaps',
	'strike',
	'doubleStrike',
	'vanish',
] as const satisfies readonly (keyof RunFormatting)[];
const FONT_SIZE_KEYS = ['fontSize', 'fontSizeComplexScript'] as const;
const OVERRIDE_KEYS = [
	'language',
	'underline',
	'underlineStyle',
	'underlineColor',
	'highlight',
	'verticalAlign',
	'fontFamily',
	'fontFamilyEastAsia',
	'fontFamilyComplexScript',
	'fontTheme',
	'color',
	'colorTheme',
	'characterSpacingTwips',
	'textScalePercent',
	'ligatures',
	'kerningHalfPoints',
	'positionHalfPoints',
	'shadingFill',
	'shadingThemeFill',
] as const satisfies readonly (keyof RunFormatting)[];

/** Every modeled direct run property, including values absent from a restored snapshot. */
export const DIRECT_RUN_PROPERTY_KEYS = [
	...TOGGLE_KEYS,
	...OVERRIDE_KEYS,
	...FONT_SIZE_KEYS,
	'eastAsiaLanguage',
	'bidiLanguage',
	'rtl',
	'style',
] as const satisfies readonly (keyof TextRun)[];

function mergeRunStyle(prior: RunFormatting, next: RunFormatting): RunFormatting {
	const result = mergeDefined(prior, next);
	replaceThemeFallbacks(result, next);
	return result;
}

function replaceThemeFallbacks(result: RunFormatting, next: RunFormatting): void {
	// A new direct value replaces the inherited reference and vice versa. If both
	// came from the same level, keep the literal as the theme reference's fallback.
	if (next.color !== undefined && next.colorTheme === undefined) delete result.colorTheme;
	if (next.colorTheme && next.color === undefined) delete result.color;
	if (next.fontFamily !== undefined && !next.fontTheme?.ascii && !next.fontTheme?.hAnsi) {
		const { ascii: _ascii, hAnsi: _hAnsi, ...scripts } = result.fontTheme ?? {};
		if (Object.keys(scripts).length) result.fontTheme = scripts;
		else delete result.fontTheme;
	}
	if ((next.fontTheme?.ascii || next.fontTheme?.hAnsi) && next.fontFamily === undefined)
		delete result.fontFamily;
}

type Resolver = (id: string | undefined) => RunFormatting | undefined;
interface ParagraphStyleRun {
	formatting: RunFormatting;
	fontSizeFromDefault: Partial<Record<(typeof FONT_SIZE_KEYS)[number], boolean>>;
}
type ParagraphRunResolver = (id: string | undefined) => ParagraphStyleRun | undefined;
const paragraphRunResolvers = new WeakMap<object, WeakMap<RunStyleCatalog, ParagraphRunResolver>>();
const characterResolvers = new WeakMap<object, Resolver>();

/** The run formatting the paragraph style chain gives (each style's `w:rPr`, nearest defined). */
function paragraphStyleRun(
	styleId: string | undefined,
	paragraphCatalog: ParagraphStyleCatalog | undefined,
	runCatalog: RunStyleCatalog | undefined,
): ParagraphStyleRun | undefined {
	if (!paragraphCatalog || !runCatalog) return undefined;
	const defaultId = defaultStyleId(paragraphCatalog.styles);
	let byRunCatalog = paragraphRunResolvers.get(paragraphCatalog.styles);
	if (!byRunCatalog) {
		byRunCatalog = new WeakMap();
		paragraphRunResolvers.set(paragraphCatalog.styles, byRunCatalog);
	}
	let resolve = byRunCatalog.get(runCatalog);
	if (!resolve) {
		resolve = chainResolver(
			paragraphCatalog.styles,
			(_style, id): ParagraphStyleRun | undefined => {
				const formatting = runCatalog.styles[id]?.formatting;
				if (!formatting) return undefined;
				const fontSizeFromDefault: ParagraphStyleRun['fontSizeFromDefault'] = {};
				for (const key of FONT_SIZE_KEYS)
					if (formatting[key] !== undefined) fontSizeFromDefault[key] = id === defaultId;
				return { formatting, fontSizeFromDefault };
			},
			(prior, next) => ({
				formatting: mergeRunStyle(prior.formatting, next.formatting),
				fontSizeFromDefault: { ...prior.fontSizeFromDefault, ...next.fontSizeFromDefault },
			}),
			{ formatting: {}, fontSizeFromDefault: {} },
		);
		byRunCatalog.set(runCatalog, resolve);
	}
	return resolve(styleId || defaultId);
}

/** The run formatting a character style chain gives, nearest defined. */
function characterStyleRun(
	styleId: string | undefined,
	runCatalog: RunStyleCatalog | undefined,
): RunFormatting {
	if (!styleId || !runCatalog) return {};
	const resolve = cachedResolver(characterResolvers, runCatalog.styles, () =>
		chainResolver(runCatalog.styles, (style) => style.formatting, mergeRunStyle, {}),
	);
	return resolve(styleId) ?? {};
}

export interface RunFormattingContext {
	// Callers forward optional lookups directly; `undefined` means "no such style/catalog".
	paragraphStyleId?: string | undefined;
	paragraphCatalog?: ParagraphStyleCatalog | undefined;
	runCatalog?: RunStyleCatalog | undefined;
	/** Simple table style run formatting for the cell/region containing this run, if any. */
	tableStyleRun?: RunFormatting;
	/** True disables the default paragraph style's table font-size exception. */
	overrideTableStyleFontSizeAndJustification?: boolean | undefined;
}

/**
 * Resolves default, style-chain and direct run formatting into one effective
 * layer for display. Direct properties on `run` and the source catalogs are
 * never mutated or flattened; this is purely a rendering-time projection.
 */
export function resolveRunFormatting(
	run: TextRun,
	context: RunFormattingContext = {},
): RunFormatting {
	const docDefaults = context.runCatalog?.docDefaults ?? {};
	const paragraph = paragraphStyleRun(
		context.paragraphStyleId,
		context.paragraphCatalog,
		context.runCatalog,
	);
	const paragraphStyle = paragraph?.formatting ?? {};
	const characterStyle = characterStyleRun(run.style, context.runCatalog);
	const tableStyle = context.tableStyleRun ?? {};
	const levels: RunFormatting[] = [docDefaults, paragraphStyle, tableStyle, characterStyle, run];
	const result: RunFormatting = {};
	// Toggle properties (ECMA-376 §17.7.3): each style type (paragraph, table, character) resolves
	// its value through its own basedOn chain; the types then combine by XOR with the document
	// default. Direct formatting on the run sets the value outright, including an explicit off.
	// Each style level already holds its chain's nearest defined value.
	const chainValue = (level: RunFormatting, key: (typeof TOGGLE_KEYS)[number]) =>
		Boolean(level[key]);
	for (const key of TOGGLE_KEYS) {
		const direct = run[key];
		const state =
			direct !== undefined
				? Boolean(direct)
				: Boolean(docDefaults[key]) !==
					(chainValue(paragraphStyle, key) !==
						(chainValue(tableStyle, key) !== chainValue(characterStyle, key)));
		if (state) result[key] = true;
	}
	for (const level of levels) {
		for (const key of OVERRIDE_KEYS) {
			const value = level[key];
			if (value !== undefined) (result as Record<string, unknown>)[key] = value;
		}
		replaceThemeFallbacks(result, level);
	}
	// MS-DOCX 2.3.1 says "specifies a font size" without distinguishing w:sz and w:szCs.
	// Apply it to each size independently when the winning paragraph declaration comes from the
	// default style, including when inherited by a child. A declaration in its base or child keeps
	// normal precedence, even if it has the same value.
	for (const key of FONT_SIZE_KEYS) {
		const paragraphSize = paragraphStyle[key];
		const tableSize = tableStyle[key];
		const useTableSize =
			tableSize !== undefined &&
			!context.overrideTableStyleFontSizeAndJustification &&
			paragraph?.fontSizeFromDefault[key] &&
			(paragraphSize === 11 || paragraphSize === 12);
		const styleSize = useTableSize ? tableSize : (paragraphSize ?? tableSize);
		const size = run[key] ?? characterStyle[key] ?? styleSize ?? docDefaults[key];
		if (size !== undefined) result[key] = size;
	}
	return result;
}
