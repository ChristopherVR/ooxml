import type { VisioPackage } from './package';
import type { VisioTextEdit } from './edit-commands';
import { replacePlainText } from './edit-text';
import { assertFormattingDependencies } from './edit-formatting-scope';

/** Plain text is not numeric: refuse affected or unknown dependencies instead of
 * retaining caches whose values would change after native text layout evaluation.
 */
export function replaceScopedPlainText(
	pkg: VisioPackage,
	pagePaths: ReadonlySet<string>,
	roots: ReadonlyMap<string, Element>,
	document: Element,
	edit: VisioTextEdit,
	check: () => void,
): Promise<boolean> {
	return replacePlainText(
		roots.get(edit.pageId)!,
		document,
		edit.shapeId,
		edit.text,
		check,
		(shape, text) =>
			assertFormattingDependencies(
				pkg,
				pagePaths,
				roots,
				edit.pageId,
				shape,
				new Map([['TheText', text]]),
				check,
			),
	);
}
