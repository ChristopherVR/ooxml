import { splitSections, tokenizeSection } from './tokenizer';

/** System placeholders stand for the long date and time formats. */
function systemFormat(source: string): string {
	if (/\[\$-(?:F800|x-sysdate)\]/i.test(source)) return 'dddd, mmmm d, yyyy';
	if (/\[\$-(?:F400|x-systime)\]/i.test(source)) return 'h:mm:ss AM/PM';
	return source;
}

/** Separate the numeric sections from the optional text section of an Office format code. */
export function numberFormatSections(format: string) {
	const sections = splitSections(format.trim() === '' ? 'General' : format)
		.slice(0, 4)
		.map((section) => tokenizeSection(systemFormat(section)));
	if (sections.length === 4) return { numeric: sections.slice(0, 3), text: sections[3] };
	const last = sections.at(-1);
	if (last?.kind === 'text') return { numeric: sections.slice(0, -1), text: last };
	return { numeric: sections, text: undefined };
}
