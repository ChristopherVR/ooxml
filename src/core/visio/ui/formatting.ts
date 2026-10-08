import type { VisioDocument, VisioPage, VisioShape } from '../model';

/** Scene-level candidate only; source formulas, rich markup and locks remain authoritative. */
export function visioStyleFormattingShape(
	page: VisioPage,
	shapeId: string,
): VisioShape | undefined {
	const candidates = page.shapes.filter((shape) => shape.id === shapeId);
	if (candidates.length !== 1) return undefined;
	const shape = candidates[0]!;
	if (
		shape.hidden ||
		shape.masterId ||
		shape.layerIds?.length ||
		shape.children.length ||
		!['shape', 'connector'].includes(shape.kind) ||
		shape.image ||
		shape.foreignVector
	)
		return undefined;
	return shape;
}
/** Text controls require uniform scene runs in addition to an eligible local leaf. */
export function visioFormattingShape(page: VisioPage, shapeId: string): VisioShape | undefined {
	const shape = visioStyleFormattingShape(page, shapeId);
	if (!shape) return undefined;
	const first = shape.text.runs[0];
	if (
		first &&
		shape.text.runs.some(
			(run) =>
				run.fontFamily !== first.fontFamily ||
				run.fontSize !== first.fontSize ||
				run.bold !== first.bold ||
				run.italic !== first.italic ||
				run.underline !== first.underline,
		)
	)
		return undefined;
	return shape;
}
/** Offer existing document fonts, avoiding invented font IDs or caches. */
export function visioFontFamilies(document: VisioDocument): readonly string[] {
	return [...new Set(document.fontFamilies ?? [])].sort((left, right) => left.localeCompare(right));
}
