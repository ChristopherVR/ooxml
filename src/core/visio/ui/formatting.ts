import type { VisioDocument, VisioPage, VisioShape, VisioText, VisioTextRun } from '../model';
import type { VisioTextFormatEdit } from '../edit-formatting-commands';

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
				run.underline !== first.underline ||
				!!run.strikethrough !== !!first.strikethrough,
		)
	)
		return undefined;
	return shape;
}
/** Offer existing document fonts, avoiding invented font IDs or caches. */
export function visioFontFamilies(document: VisioDocument): readonly string[] {
	return [...new Set(document.fontFamilies ?? [])].sort((left, right) => left.localeCompare(right));
}

export interface VisioTextFormattingState {
	bold: boolean;
	italic: boolean;
	underline: boolean;
	strikethrough: boolean;
	bullets: boolean;
	fontFamily: string | undefined;
	fontSize: number | undefined;
	fontColor: string | undefined;
	horizontalAlign: VisioTextFormatEdit['horizontalAlign'];
	verticalAlign: VisioTextFormatEdit['verticalAlign'];
	canIndentDecrease: boolean;
}
/** Mixed selection styles toggle on unless every selected run already has that style. */
export function visioTextFormattingState(shapes: readonly VisioShape[]): VisioTextFormattingState {
	const common = <T>(values: readonly T[]): T | undefined =>
		values.length && values.every((value) => value === values[0]) ? values[0] : undefined;
	const runs = shapes.flatMap<VisioText | VisioTextRun>((shape) =>
		shape.text.runs.length ? shape.text.runs : [shape.text],
	);
	const styled = (key: 'bold' | 'italic' | 'underline' | 'strikethrough') =>
		!!runs.length && runs.every((run) => !!run[key]);
	const paragraphs = shapes.flatMap((shape) => shape.text.paragraphs ?? []);
	const alignment = common(
		shapes.flatMap((shape) =>
			shape.text.paragraphs?.length
				? shape.text.paragraphs.map((paragraph) => paragraph.horizontalAlign)
				: [shape.text.horizontalAlign],
		),
	);
	return {
		bold: styled('bold'),
		italic: styled('italic'),
		underline: styled('underline'),
		strikethrough: styled('strikethrough'),
		bullets:
			!!shapes.length &&
			shapes.every(
				(shape) =>
					!!shape.text.paragraphs?.length &&
					shape.text.paragraphs.every((paragraph) => !!paragraph.bullet),
			),
		fontFamily: common(runs.map((run) => run.fontFamily)),
		fontSize: common(runs.map((run) => run.fontSize)),
		fontColor: common(runs.map((run) => run.color)),
		horizontalAlign: alignment === 'distributed' ? undefined : alignment,
		verticalAlign: common(shapes.map((shape) => shape.text.verticalAlign)),
		canIndentDecrease: paragraphs.some((paragraph) => paragraph.indentLeft > 0),
	};
}

/** Ribbon indent increments use 18 physical points and do not follow page drawing scale. */
export function visioTextIndentCommand(
	page: VisioPage,
	shapeId: string,
	direction: 'increase' | 'decrease',
): VisioTextFormatEdit | undefined {
	const shape = visioFormattingShape(page, shapeId);
	if (!shape || !['increase', 'decrease'].includes(direction)) return undefined;
	const paragraphs = shape.text.paragraphs ?? [];
	const left = paragraphs[0]?.indentLeft ?? 0;
	if (
		!Number.isFinite(left) ||
		left < 0 ||
		paragraphs.some((paragraph) => paragraph.indentLeft !== left)
	)
		return undefined;
	const points = left * 72;
	const indentLeft = Math.min(7200, Math.max(0, points + (direction === 'increase' ? 18 : -18)));
	return { type: 'format-text', pageId: page.id, shapeId, indentLeft };
}
