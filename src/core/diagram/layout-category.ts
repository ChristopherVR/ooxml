// Resolves a layout family from a layout definition's unique id and gallery categories.
// Extracted from `pptx/core/core/runtime/smartart-layout-category.ts`.
import type { DiagramLayoutType } from './types.js';

/** Canonical `dgm:cat/@type` values that map 1:1 onto layout families. */
const CATEGORY_FAMILY: Record<string, DiagramLayoutType> = {
	list: 'list',
	process: 'process',
	cycle: 'cycle',
	hierarchy: 'hierarchy',
	relationship: 'relationship',
	matrix: 'matrix',
	pyramid: 'pyramid',
};

const UNIQUE_ID_KEYWORDS: Array<[RegExp, DiagramLayoutType]> = [
	[/hier|org/u, 'hierarchy'],
	[/cycle|radial|circular/u, 'cycle'],
	[/process|flow|chevron|arrow|equation/u, 'process'],
	[/timeline/u, 'timeline'],
	[/matrix|grid/u, 'matrix'],
	[/pyramid/u, 'pyramid'],
	[/venn/u, 'venn'],
	[/funnel/u, 'funnel'],
	[/gear/u, 'gear'],
	[/target/u, 'target'],
	[/list/u, 'list'],
];

/**
 * The layout family of a `dgm:layoutDef`: the catalogue category when it names one, else a
 * keyword of the unique id. `undefined` when neither is decisive, so callers can leave the family
 * unset rather than guess.
 */
export function resolveDiagramLayoutCategory(
	uniqueId: string,
	categories: readonly string[],
): DiagramLayoutType | undefined {
	for (const category of categories) {
		const family = CATEGORY_FAMILY[category.trim().toLowerCase()];
		if (family) return family;
	}
	const lowerId = uniqueId.toLowerCase();
	if (lowerId.length > 0)
		for (const [pattern, family] of UNIQUE_ID_KEYWORDS) if (pattern.test(lowerId)) return family;
	return undefined;
}
