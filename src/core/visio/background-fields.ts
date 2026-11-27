import type { VisioPage, VisioShape, VisioText } from './model';

/** The foreground page a background page is drawn under. */
export interface VisioForegroundPage {
	name: string;
	/** 1-based among the foreground pages; omitted for a background page. */
	number?: number;
}

const PAGE_NAME = /^=?\s*PAGENAME\s*\(\s*0?\s*\)\s*$/i;
const PAGE_NUMBER = /^=?\s*PAGENUMBER\s*\(\s*\)\s*$/i;

/**
 * The shapes of a background page as Visio draws them under `foreground`: a page field such as
 * PAGENAME() or PAGENUMBER() on a background page shows the foreground page that uses it, not the
 * background's own name. Text is parsed once per page, so the background's display text is
 * rewritten here; shapes without such a field are returned as they are, and `onCopy` hears of
 * every shape that is replaced so per-shape lookups (visibility) can follow it.
 */
export function visioBackgroundFieldShapes(
	background: VisioPage,
	foreground: VisioForegroundPage,
	onCopy?: (source: VisioShape, copy: VisioShape) => void,
): VisioShape[] {
	const visit = (shapes: VisioShape[]): VisioShape[] => {
		let changed = false;
		const result = shapes.map((shape) => {
			const text = shape.text && retargetText(shape.text, background.name, foreground);
			const children = shape.children.length ? visit(shape.children) : shape.children;
			if (!text && children === shape.children) return shape;
			changed = true;
			const copy: VisioShape = { ...shape, children, ...(text ? { text } : {}) };
			onCopy?.(shape, copy);
			return copy;
		});
		return changed ? result : shapes;
	};
	return visit(background.shapes);
}

/** The page-field display of one field, or undefined to keep it. */
function display(
	formula: string | undefined,
	shown: string,
	own: string,
	foreground: VisioForegroundPage,
): string | undefined {
	if (!formula) return undefined;
	if (PAGE_NAME.test(formula)) {
		// The field showed the background's own name, perhaps through an upper or lower case picture.
		if (shown === own) return foreground.name;
		if (shown === own.toUpperCase()) return foreground.name.toUpperCase();
		if (shown === own.toLowerCase()) return foreground.name.toLowerCase();
		return undefined;
	}
	if (PAGE_NUMBER.test(formula) && foreground.number !== undefined && /^\d+$/.test(shown))
		return String(foreground.number);
	return undefined;
}

/** The text with its page fields rewritten, or undefined when none changes. */
function retargetText(
	text: VisioText,
	own: string,
	foreground: VisioForegroundPage,
): VisioText | undefined {
	if (!text.fields?.length) return undefined;
	const edits: { start: number; end: number; value: string }[] = [];
	for (const field of text.fields) {
		const shown = text.plainText.slice(field.start, field.end);
		const value = display(field.formula, shown, own, foreground);
		if (value !== undefined && value !== shown) edits.push({ ...field, value });
	}
	if (!edits.length) return undefined;
	// Later edits first, so earlier offsets stay valid while splicing.
	edits.sort((a, b) => b.start - a.start);
	let plainText = text.plainText;
	let runs = text.runs;
	for (const edit of edits) {
		plainText = plainText.slice(0, edit.start) + edit.value + plainText.slice(edit.end);
		runs = spliceRuns(runs, edit.start, edit.end, edit.value);
	}
	// An offset moves by every edit that ends before it; an edit ending exactly at it moves an end
	// offset, and a start offset only when the edit is not an insertion there.
	const shift = (offset: number, isEnd: boolean) => {
		let moved = offset;
		for (const edit of edits)
			if (edit.end < offset || (edit.end === offset && (isEnd || edit.start < edit.end)))
				moved += edit.value.length - (edit.end - edit.start);
		return moved;
	};
	const fields = text.fields.map((field) => {
		const edit = edits.find((candidate) => candidate.start === field.start);
		const start = shift(field.start, false);
		return { ...field, start, end: edit ? start + edit.value.length : shift(field.end, true) };
	});
	return {
		...text,
		plainText,
		runs,
		fields,
		...(text.paragraphs
			? {
					paragraphs: text.paragraphs.map((paragraph) => ({
						...paragraph,
						start: shift(paragraph.start, false),
						end: shift(paragraph.end, true),
					})),
				}
			: {}),
	};
}

/** Replace the characters [start, end) across the runs, keeping the first run's formatting. */
function spliceRuns(
	runs: VisioText['runs'],
	start: number,
	end: number,
	value: string,
): VisioText['runs'] {
	let offset = 0;
	let placed = false;
	const result: VisioText['runs'] = [];
	for (const run of runs) {
		const from = offset;
		const to = offset + run.text.length;
		offset = to;
		if (to <= start || from >= end) {
			// A run wholly before or after the field, or an empty field at a run boundary.
			if (!placed && from >= end && from === start) {
				result.push({ ...run, text: value + run.text });
				placed = true;
				continue;
			}
			result.push(run);
			continue;
		}
		const before = run.text.slice(0, Math.max(0, start - from));
		const after = run.text.slice(Math.max(0, end - from));
		const text = before + (placed ? '' : value) + after;
		placed = true;
		if (text) result.push({ ...run, text });
	}
	if (!placed && result.length) {
		const last = result[result.length - 1]!;
		result[result.length - 1] = { ...last, text: last.text + value };
	}
	return result;
}
