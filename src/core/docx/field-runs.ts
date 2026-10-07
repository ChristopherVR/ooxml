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
 * Follows `w:fldChar` begin/separate/end markers across a paragraph's runs. Runs between
 * `separate` and `end` receive the innermost field's cached-result metadata. The tracker
 * lifetime remains paragraph-local; fields spanning paragraphs need a container-level tracker.
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
