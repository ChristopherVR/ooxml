import type { TextRun } from './model';
import type { RunFormattingContext } from './run-formatting';
import type { RunStyleCatalog } from './run-style-model';
import { cachedResolver, chainResolver, defaultStyleId } from './style-inheritance';

const languageResolvers = new WeakMap<
	object,
	(id: string | undefined) => { eastAsiaLanguage?: string } | undefined
>();

/**
 * The East Asian language of a run (`w:lang/@w:eastAsia`) through the style hierarchy: direct,
 * then the character style chain, the paragraph style chain and the document defaults. Like the
 * rest of run formatting this is a display projection; nothing is flattened into the run.
 */
export function resolveEastAsiaLanguage(
	run: TextRun,
	context: RunFormattingContext = {},
): string | undefined {
	if (run.eastAsiaLanguage !== undefined) return run.eastAsiaLanguage;
	const catalog = context.runCatalog;
	if (!catalog) return undefined;
	// The nearest style of the chain that names a language.
	const resolve = cachedResolver(languageResolvers, catalog.styles, () =>
		chainResolver<RunStyleCatalog['styles'][string], { eastAsiaLanguage?: string }>(
			catalog.styles,
			(style) =>
				style.eastAsiaLanguage === undefined
					? undefined
					: { eastAsiaLanguage: style.eastAsiaLanguage },
			(_resolved, nearer) => nearer,
			{},
		),
	);
	const chainLanguage = (styleId: string | undefined) =>
		resolve(styleId || undefined)?.eastAsiaLanguage;
	const paragraphStyle =
		context.paragraphStyleId ||
		(context.paragraphCatalog ? defaultStyleId(context.paragraphCatalog.styles) : undefined);
	return (
		chainLanguage(run.style) ?? chainLanguage(paragraphStyle) ?? catalog.docDefaultsEastAsiaLanguage
	);
}
