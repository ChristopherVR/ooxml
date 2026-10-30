// The Replace with text of the Find and Replace dialog, expanded for one match: the found text
// (`^&`), the groups of a wildcard pattern (`\1` to `\9`) and the Find codes that stand for breaks,
// tabs and special characters.
import { FindPatternError } from './find-syntax';

export type ReplacementBreak =
	| 'paragraph'
	| 'tab'
	| 'lineBreak'
	| 'pageBreak'
	| 'optionalHyphen'
	| 'nonBreakingHyphen'
	| 'nonBreakingSpace';

export type ReplacementPart =
	| { readonly kind: 'text'; readonly text: string }
	| { readonly kind: ReplacementBreak };

const BREAK_CODES: Readonly<Record<string, ReplacementBreak>> = {
	p: 'paragraph',
	t: 'tab',
	l: 'lineBreak',
	m: 'pageBreak',
	'-': 'optionalHyphen',
	'~': 'nonBreakingHyphen',
	s: 'nonBreakingSpace',
};

/**
 * The parts of a replacement for one match. `wildcards` makes `\1`..`\9` the pattern's groups;
 * without it a backslash is an ordinary character. A typed line break becomes a line break.
 */
export function replacementParts(
	template: string,
	match: { readonly text: string; readonly groups: readonly string[] },
	wildcards = false,
): ReplacementPart[] {
	const parts: ReplacementPart[] = [];
	const push = (part: ReplacementPart): void => {
		const last = parts[parts.length - 1];
		if (part.kind === 'text' && last?.kind === 'text')
			parts[parts.length - 1] = { kind: 'text', text: last.text + part.text };
		else parts.push(part);
	};
	const text = (value: string): void => push({ kind: 'text', text: value });
	for (let index = 0; index < template.length; index++) {
		const char = template.charAt(index);
		if (char === '^') {
			const code = template.charAt(index + 1);
			if (!code) throw new FindPatternError('The replacement ends with ^');
			index++;
			if (code === '&') text(match.text);
			else if (code === '^') text('^');
			else {
				const kind = BREAK_CODES[code.toLowerCase()];
				if (!kind) throw new FindPatternError(`Unknown replacement code ^${code}`);
				push({ kind });
			}
		} else if (char === '\\' && wildcards) {
			const next = template.charAt(index + 1);
			if (!next) throw new FindPatternError('The replacement ends with \\');
			index++;
			if (/^[1-9]$/.test(next)) {
				const group = match.groups[Number(next)];
				if (group === undefined) throw new FindPatternError(`The pattern has no group ${next}`);
				text(group);
			} else text(next);
		} else if (char === '\r' || char === '\n') {
			if (char === '\r' && template.charAt(index + 1) === '\n') index++;
			push({ kind: 'lineBreak' });
		} else {
			text(char);
		}
	}
	return parts;
}
