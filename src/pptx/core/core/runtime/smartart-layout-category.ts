/**
 * Resolve a {@link SmartArtLayoutType} family from a diagram LAYOUT part
 * (`dgm:layoutDef`). The part filename (`layout1`) says nothing about the
 * algorithm, so the loader previously left `resolvedLayoutType` unset and
 * renderers fell back to a plain list for any diagram without a cached
 * drawing part. The layout definition itself carries the answer twice:
 * `dgm:catLst/dgm:cat/@type` (canonical gallery categories) and the
 * `@uniqueId` URN (e.g. `urn:microsoft.com/office/officeart/2005/8/layout/orgChart1`).
 */
import { resolveDiagramLayoutCategory } from '../../../../diagram/index.js';
import type { SmartArtLayoutType } from '../../types';

/**
 * Resolve the layout family from a layout definition's unique id and its
 * gallery categories. Returns `undefined` when neither signal is decisive so
 * callers can leave `resolvedLayoutType` unset rather than guess. The logic lives in the
 * format-neutral `diagram` area.
 */
export function resolveSmartArtLayoutCategory(
	uniqueId: string,
	categories: readonly string[],
): SmartArtLayoutType | undefined {
	return resolveDiagramLayoutCategory(uniqueId, categories);
}
