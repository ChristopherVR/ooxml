// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
import type { ParagraphStyleCatalog, TextRun } from './model';
import type { RunFormatting, RunStyleCatalog } from './run-style-model';
import { cachedResolver, chainResolver, defaultStyleId, mergeDefined } from './style-inheritance';

/**
 * Word "toggle" run properties (ECMA-376 17.7.3): each level in the formatting
 * hierarchy that specifies the property flips its effective state, regardless
 * of the level's own boolean value — so two ancestor levels that both turn
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
const OVERRIDE_KEYS = [
	'underline',
	'underlineStyle',
	'underlineColor',
	'highlight',
	'verticalAlign',
	'fontSize',
	'fontFamily',
	'fontFamilyEastAsia',
	'fontFamilyComplexScript',
	'fontSizeComplexScript',
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
	'language',
	'eastAsiaLanguage',
	'bidiLanguage',
	'rtl',
	'style',
] as const satisfies readonly (keyof TextRun)[];

type Resolver = (id: string | undefined) => RunFormatting | undefined;
const paragraphRunResolvers = new WeakMap<object, WeakMap<RunStyleCatalog, Resolver>>();
const characterResolvers = new WeakMap<object, Resolver>();

/** The run formatting the paragraph style chain gives (each style's `w:rPr`, nearest defined). */
function paragraphStyleRun(
	styleId: string | undefined,
	paragraphCatalog: ParagraphStyleCatalog | undefined,
	runCatalog: RunStyleCatalog | undefined,
): RunFormatting {
	if (!paragraphCatalog || !runCatalog) return {};
	let byRunCatalog = paragraphRunResolvers.get(paragraphCatalog.styles);
	if (!byRunCatalog) {
		byRunCatalog = new WeakMap();
		paragraphRunResolvers.set(paragraphCatalog.styles, byRunCatalog);
	}
	let resolve = byRunCatalog.get(runCatalog);
	if (!resolve) {
		resolve = chainResolver(
			paragraphCatalog.styles,
			(_style, id) => runCatalog.styles[id]?.formatting,
			mergeDefined,
			{},
		);
		byRunCatalog.set(runCatalog, resolve);
	}
	return resolve(styleId || defaultStyleId(paragraphCatalog.styles)) ?? {};
}

/** The run formatting a character style chain gives, nearest defined. */
function characterStyleRun(
	styleId: string | undefined,
	runCatalog: RunStyleCatalog | undefined,
): RunFormatting {
	if (!styleId || !runCatalog) return {};
	const resolve = cachedResolver(characterResolvers, runCatalog.styles, () =>
		chainResolver(runCatalog.styles, (style) => style.formatting, mergeDefined, {}),
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
	const paragraphStyle = paragraphStyleRun(
		context.paragraphStyleId,
		context.paragraphCatalog,
		context.runCatalog,
	);
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
	for (const key of OVERRIDE_KEYS) {
		for (const level of levels) {
			const value = level[key];
			if (value !== undefined) (result as Record<string, unknown>)[key] = value;
		}
	}
	return result;
}
