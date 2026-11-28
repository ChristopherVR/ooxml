// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Tags field result runs with their field code so renderers can recalculate display-only fields
// such as PAGE and NUMPAGES. The markers and field code are modeled as their own runs.
import type { TextRun } from './model';
import { getW, isElement, named, textContent, type XmlElement } from './xml';
import { parseFieldFlags } from './field-flags';
import { isCachedPageBreak } from './breaks';

const isInstruction = (element: XmlElement) =>
	named(element, 'instrText') || named(element, 'delInstrText');

/** Word can pair a carriage return with instruction text in the same run. */
export function runFieldInstruction(content: readonly XmlElement[]): string | undefined {
	if (!content.some(isInstruction)) return undefined;
	let instruction = '';
	for (const child of content) {
		if (isInstruction(child)) instruction += textContent(child);
		else if (named(child, 'cr')) instruction += '\r';
		else if (!named(child, 'rPr') && !isCachedPageBreak(child)) return undefined;
	}
	return instruction;
}

/** The field code's name (first word, upper case), e.g. `PAGE` for ` PAGE \* MERGEFORMAT `. */
export function fieldName(instr: string): string {
	return instr.trim().split(/\s+/)[0]?.toUpperCase() ?? '';
}

/**
 * Follows `w:fldChar` begin/separate/end markers across runs. Runs between `separate` and
 * `end` receive the innermost field's cached-result metadata. Share a tracker within one story.
 */
export function createFieldTracker() {
	const stack: { instr: string; inResult: boolean; locked?: boolean }[] = [];
	return (element: XmlElement, run: TextRun): TextRun => {
		const children = Array.from(element.childNodes).filter(isElement);
		const open = stack.at(-1);
		if (open && !open.inResult) {
			const instruction = run.fieldCode ?? runFieldInstruction(children);
			if (instruction !== undefined) {
				open.instr += instruction;
				return run;
			}
		}
		let resultInstr: string | undefined;
		let resultLocked: boolean | undefined;
		for (const child of children) {
			if (named(child, 'fldChar')) {
				const type = getW(child, 'fldCharType');
				if (type === 'begin') {
					const locked = parseFieldFlags(child)?.locked;
					stack.push({ instr: '', inResult: false, ...(locked !== undefined && { locked }) });
				} else if (type === 'separate') {
					const top = stack.at(-1);
					if (top) top.inResult = true;
				} else if (type === 'end') stack.pop();
			} else {
				const top = stack.at(-1);
				if (isInstruction(child) && top && !top.inResult) top.instr += textContent(child);
				else if ((named(child, 't') || named(child, 'delText')) && top?.inResult) {
					resultInstr = top.instr.trim();
					resultLocked = top.locked;
				}
			}
		}
		if (resultInstr) {
			run.field = { instr: resultInstr };
			// The cache mirrors lock state for recalculation; dirty remains on the structural marker.
			if (resultLocked !== undefined) run.fieldFlags = { locked: resultLocked };
		}
		return run;
	};
}

export type FieldResultMetadata = WeakMap<
	XmlElement,
	{ field: NonNullable<TextRun['field']>; fieldFlags?: TextRun['fieldFlags'] }
>;

/**
 * Scan one story in XML order, including read-only nested table content that the model omits.
 * Runs are terminal so drawing/textbox stories cannot affect this story's stack. Simple fields
 * own their cached content and cannot modify an enclosing complex field's tracking state.
 */
export function storyFieldResults(container: XmlElement): FieldResultMetadata {
	const metadata: FieldResultMetadata = new WeakMap();
	const trackField = createFieldTracker();
	const visit = (element: XmlElement) => {
		// A text box's own story is scanned when it is parsed as the container.
		if (named(element, 'fldSimple') || (element !== container && named(element, 'txbxContent')))
			return;
		if (named(element, 'r')) {
			const run = trackField(element, { text: '' });
			if (run.field)
				metadata.set(element, {
					field: run.field,
					...(run.fieldFlags && { fieldFlags: run.fieldFlags }),
				});
			return;
		}
		for (const child of Array.from(element.childNodes)) if (isElement(child)) visit(child);
	};
	visit(container);
	return metadata;
}
