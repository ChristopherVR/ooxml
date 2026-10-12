import type { VisioDocument, VisioShape } from './model';

export interface VisioAccessibilityIssue {
	pageId: string;
	shapeId: string;
	shapeName: string;
	code: 'missing-alternative-text';
}

/** Check unlabeled visible graphics. Contrast and reading order require separate review. */
export function checkVisioAlternativeText(document: VisioDocument): VisioAccessibilityIssue[] {
	const issues: VisioAccessibilityIssue[] = [];
	for (const page of document.pages) {
		const visit = (shapes: readonly VisioShape[]): void => {
			for (const shape of shapes) {
				if (shape.hidden) continue;
				const named = [shape.accessibilityTitle, shape.alternativeText, shape.text.plainText].some(
					(value) => value?.trim(),
				);
				if (!named && shape.kind !== 'group')
					issues.push({
						pageId: page.id,
						shapeId: shape.id,
						shapeName: shape.name,
						code: 'missing-alternative-text',
					});
				// A described group provides a single alternative for its constituent graphics.
				if (!named) visit(shape.children);
			}
		};
		visit(page.shapes);
	}
	return issues;
}
