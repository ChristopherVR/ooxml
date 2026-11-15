import type { TextRun } from './model';
import type { RunFormattingContext } from './run-formatting';

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
	const chainLanguage = (styleId: string | undefined) => {
		const visited = new Set<string>();
		for (let current = styleId; current && !visited.has(current);) {
			visited.add(current);
			const style = catalog.styles[current];
			if (!style) return undefined;
			if (style.eastAsiaLanguage !== undefined) return style.eastAsiaLanguage;
			current = style.basedOn;
		}
		return undefined;
	};
	const paragraphStyle =
		context.paragraphStyleId ||
		(context.paragraphCatalog
			? Object.values(context.paragraphCatalog.styles).find((style) => style.isDefault)?.id
			: undefined);
	return (
		chainLanguage(run.style) ?? chainLanguage(paragraphStyle) ?? catalog.docDefaultsEastAsiaLanguage
	);
}
