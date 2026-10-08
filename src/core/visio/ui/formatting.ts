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
/** Mixed runs can be formatted together; source row admission remains authoritative. */
export function visioFormattingShape(page: VisioPage, shapeId: string): VisioShape | undefined {
	return visioStyleFormattingShape(page, shapeId);
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

/** A size step requires a common rendered run size; absolute size changes admit mixed runs. */
export function visioTextFontStepCommand(
	page: VisioPage,
	shapeId: string,
	direction: 'increase' | 'decrease',
): VisioTextFormatEdit | undefined {
	const shape = visioFormattingShape(page, shapeId);
	if (!shape || !['increase', 'decrease'].includes(direction)) return undefined;
	const sizes = shape.text.runs.length
		? shape.text.runs.map((run) => run.fontSize)
		: [shape.text.fontSize];
	const size = sizes[0]!;
	if (!Number.isFinite(size) || sizes.some((value) => Math.abs(value - size) > 1e-10))
		return undefined;
	const current = size * 72;
	if (current < 1 || current > 1000) return undefined;
	const steps = [6, 8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72];
	const fontSize =
		direction === 'increase'
			? (steps.find((value) => value > current + 0.001) ?? Math.min(1000, Math.ceil(current * 1.2)))
			: (steps.reverse().find((value) => value < current - 0.001) ??
				Math.max(1, Math.floor(current / 1.2)));
	return { type: 'format-text', pageId: page.id, shapeId, fontSize };
}
