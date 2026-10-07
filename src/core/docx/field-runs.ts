// Canonical modern DOCX implementation; legacy CFB codecs live in ole2.
// Tags field result runs with their field code so renderers can recalculate display-only fields
// such as PAGE and NUMPAGES. The markers and field code are modeled as their own runs.
import type { TextRun } from './model';
import { getW, isElement, named, textContent, type XmlElement } from './xml';

/** The field code's name (first word, upper case), e.g. `PAGE` for ` PAGE \* MERGEFORMAT `. */
export function fieldName(instr: string): string {
	return instr.trim().split(/\s+/)[0]?.toUpperCase() ?? '';
}

/**
 * Follows `w:fldChar` begin/separate/end markers across runs. Runs between `separate` and
 * `end` receive the innermost field's cached-result metadata. Share a tracker within one story.
 */
export function createFieldTracker() {
	const stack: { instr: string; inResult: boolean }[] = [];
	return (element: XmlElement, run: TextRun): TextRun => {
		let resultInstr: string | undefined;
		for (const child of Array.from(element.childNodes)) {
			if (!isElement(child)) continue;
			if (named(child, 'fldChar')) {
				const type = getW(child, 'fldCharType');
				if (type === 'begin') stack.push({ instr: '', inResult: false });
				else if (type === 'separate') {
					const top = stack.at(-1);
					if (top) top.inResult = true;
				} else if (type === 'end') stack.pop();
			} else {
				const top = stack.at(-1);
				if (named(child, 'instrText') && top && !top.inResult) top.instr += textContent(child);
				else if ((named(child, 't') || named(child, 'delText')) && top?.inResult)
					resultInstr = top.instr.trim();
			}
		}
		if (resultInstr) run.field = { instr: resultInstr };
		return run;
	};
}

export type FieldResultMetadata = WeakMap<XmlElement, NonNullable<TextRun['field']>>;

/**
 * Scan one story in XML order, including read-only nested table content that the model omits.
 * Runs are terminal so drawing/textbox stories cannot affect this story's stack. Simple fields
 * own their cached content and cannot modify an enclosing complex field's tracking state.
 */
export function storyFieldResults(container: XmlElement): FieldResultMetadata {
	const metadata: FieldResultMetadata = new WeakMap();
	const trackField = createFieldTracker();
	const visit = (element: XmlElement) => {
		if (named(element, 'fldSimple') || named(element, 'txbxContent')) return;
		if (named(element, 'r')) {
			const run = trackField(element, { text: '' });
			if (run.field) metadata.set(element, run.field);
			return;
		}
		for (const child of Array.from(element.childNodes)) if (isElement(child)) visit(child);
	};
	visit(container);
	return metadata;
}
