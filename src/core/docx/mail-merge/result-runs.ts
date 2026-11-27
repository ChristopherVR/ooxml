// The runs a merged field leaves: its result alone (unlinked), or the field with a new cached
// result, formatted as Word formats a field result.
import {
	replaceFieldResult,
	withFieldCodeFormatting,
	type FieldResultSpan,
} from '../field-result-spans';
import type { Paragraph, TextRun } from '../model';

/**
 * The run whose formatting a new result takes, as Word gives it: `\* CHARFORMAT` the first
 * letter of the field name, `\* MERGEFORMAT` the old result (undefined here), and without
 * either the last run of the field's own code (measured with Word's Update Field).
 */
function formatRun(runs: readonly TextRun[], span: FieldResultSpan, charFormat: boolean) {
	if (charFormat) return span.name ?? span.code;
	if (MERGEFORMAT.test(span.instr) || span.code === undefined) return undefined;
	let depth = 0;
	let last: number | undefined;
	for (let at = span.from + 1; at < span.insertAt; at++) {
		const run = runs[at]!;
		if (run.fieldChar === 'begin') depth++;
		else if (run.fieldChar === 'end') depth--;
		else if (!depth && run.fieldCode !== undefined && run.fieldCode.trim()) last = at;
	}
	return last;
}
const MERGEFORMAT = /\\[*]\s*MERGEFORMAT\b/i;

export function resultRun(
	paragraph: Paragraph,
	span: FieldResultSpan,
	text: string,
	charFormat: boolean,
): TextRun {
	const at = formatRun(paragraph.runs, span, charFormat);
	const code = at === undefined ? undefined : paragraph.runs[at];
	const fallback = paragraph.runs[span.code ?? span.from];
	const source = paragraph.runs[span.runs[0] ?? -1] ?? fallback ?? { text: '' };
	const formatted = code ? withFieldCodeFormatting(source, code, text) : source;
	const {
		field: _field,
		fieldInstanceId: _fieldInstanceId,
		fieldChar: _fieldChar,
		fieldCode: _fieldCode,
		fieldFlags: _fieldFlags,
		...run
	} = formatted;
	return { ...run, text: text.replace(/\r\n?/g, '\n') };
}

/** The runs of `span` with its result set to `result`, the field kept around it. */
export function keptField(
	paragraph: Paragraph,
	span: FieldResultSpan,
	result: { text: string },
	charFormat: boolean,
): TextRun[] {
	const runs = paragraph.runs.slice(span.from, span.to + 1);
	const at = (index: number | undefined) => (index === undefined ? undefined : index - span.from);
	const local: FieldResultSpan = {
		...span,
		from: 0,
		to: span.to - span.from,
		runs: span.runs.map((index) => index - span.from),
		insertAt: span.insertAt - span.from,
		...(span.code !== undefined && { code: at(span.code)! }),
		...(span.name !== undefined && { name: at(span.name)! }),
	};
	// A nested field in the old result goes with it: the new result is plain text.
	const nested = new Set<number>();
	let depth = 0;
	runs.forEach((run, index) => {
		if (index <= local.insertAt - 1 || index === local.to) return;
		if (run.fieldChar === 'begin') depth++;
		if (depth) nested.add(index);
		if (run.fieldChar === 'end' && depth) depth--;
	});
	const plain = nested.size ? runs.filter((_, index) => !nested.has(index)) : runs;
	if (nested.size) {
		const shift = (index: number) => index - [...nested].filter((n) => n < index).length;
		local.runs = local.runs.filter((index) => !nested.has(index)).map(shift);
		local.to = shift(local.to);
	}
	const text = result.text.replace(/\r\n?/g, '\n');
	const source = formatRun(plain, local, charFormat);
	return (
		replaceFieldResult(
			plain,
			source === undefined ? local : { ...local, name: source },
			text,
			source !== undefined,
		) ?? plain
	);
}
