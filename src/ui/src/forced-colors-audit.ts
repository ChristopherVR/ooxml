/**
 * Helpers for the forced-colours stylesheet audit (`forced-colors.test.ts`): CSS specificity and
 * the CSS a module keeps in template strings. Test-only; nothing in the package imports this.
 */

export type Specificity = [ids: number, classes: number, elements: number];

const add = (a: Specificity, b: Specificity): Specificity => [
	a[0] + b[0],
	a[1] + b[1],
	a[2] + b[2],
];
const greater = (a: Specificity, b: Specificity) =>
	a[0] !== b[0] ? a[0] > b[0] : a[1] !== b[1] ? a[1] > b[1] : a[2] > b[2];
const max = (selectors: Specificity[]): Specificity =>
	selectors.reduce<Specificity>((best, next) => (greater(next, best) ? next : best), [0, 0, 0]);

/** The index after the bracket that closes the one at `start`. */
function closing(text: string, start: number): number {
	const open = text[start];
	const close = open === '(' ? ')' : ']';
	let depth = 0;
	for (let i = start; i < text.length; i++) {
		if (text[i] === open) depth++;
		else if (text[i] === close && --depth === 0) return i + 1;
	}
	return text.length;
}

/** Split a selector list on its top-level commas. */
function list(text: string): string[] {
	const out: string[] = [];
	let depth = 0;
	let start = 0;
	for (let i = 0; i < text.length; i++) {
		if (text[i] === '(' || text[i] === '[') depth++;
		else if (text[i] === ')' || text[i] === ']') depth--;
		else if (text[i] === ',' && !depth) {
			out.push(text.slice(start, i));
			start = i + 1;
		}
	}
	out.push(text.slice(start));
	return out;
}

const LEGACY_PSEUDO_ELEMENTS = new Set(['before', 'after', 'first-line', 'first-letter']);

/** CSS specificity of one complex selector, following Selectors Level 4. */
export function specificity(selector: string): Specificity {
	let total: Specificity = [0, 0, 0];
	let i = 0;
	const ident = () => {
		const match = /^-?[\w-]+/.exec(selector.slice(i));
		i += match ? match[0].length : 0;
		return match ? match[0] : '';
	};
	const argument = () => {
		if (selector[i] !== '(') return '';
		const end = closing(selector, i);
		const inner = selector.slice(i + 1, end - 1);
		i = end;
		return inner;
	};
	while (i < selector.length) {
		const char = selector[i]!;
		if (char === '#') {
			i++;
			ident();
			total = add(total, [1, 0, 0]);
		} else if (char === '.') {
			i++;
			ident();
			total = add(total, [0, 1, 0]);
		} else if (char === '[') {
			i = closing(selector, i);
			total = add(total, [0, 1, 0]);
		} else if (char === ':' && selector[i + 1] === ':') {
			i += 2;
			ident();
			const inner = argument();
			total = add(total, [0, 0, 1]);
			if (inner) total = add(total, max(list(inner).map(specificity)));
		} else if (char === ':') {
			i++;
			const name = ident().toLowerCase();
			const inner = argument();
			if (name === 'where') continue;
			if (LEGACY_PSEUDO_ELEMENTS.has(name)) total = add(total, [0, 0, 1]);
			else if (name === 'is' || name === 'not' || name === 'has' || name === 'matches')
				total = add(total, max(list(inner).map(specificity)));
			else if (name === 'host' || name === 'host-context')
				total = add(add(total, [0, 1, 0]), inner ? max(list(inner).map(specificity)) : [0, 0, 0]);
			else total = add(total, [0, 1, 0]);
		} else if (/[a-z]/i.test(char)) {
			ident();
			total = add(total, [0, 0, 1]);
		} else i++;
	}
	return total;
}

/**
 * The text of every template literal in a module that reads as CSS (a `selector { property: ... }`
 * block), with each `${...}` replaced by `x`. Comments and ordinary strings are skipped.
 */
export function cssInTypeScript(source: string): string {
	const templates: string[] = [];
	let i = 0;
	const skipQuoted = (quote: string) => {
		i++;
		while (i < source.length && source[i] !== quote) i += source[i] === '\\' ? 2 : 1;
		i++;
	};
	const readTemplate = () => {
		let text = '';
		i++;
		while (i < source.length && source[i] !== '`') {
			if (source[i] === '\\') {
				text += source.slice(i, i + 2);
				i += 2;
			} else if (source[i] === '$' && source[i + 1] === '{') {
				i += 2;
				code(1);
				text += 'x';
			} else text += source[i++];
		}
		i++;
		templates.push(text);
	};
	// Walks code until the brace that closes `depth` (or the end of the module).
	function code(depth: number) {
		while (i < source.length && depth) {
			const char = source[i]!;
			if (char === '{') {
				depth++;
				i++;
			} else if (char === '}') {
				depth--;
				i++;
			} else if (char === '`') readTemplate();
			else if (char === "'" || char === '"') skipQuoted(char);
			else if (char === '/' && source[i + 1] === '/') {
				const end = source.indexOf('\n', i);
				i = end < 0 ? source.length : end;
			} else if (char === '/' && source[i + 1] === '*') {
				const end = source.indexOf('*/', i + 2);
				i = end < 0 ? source.length : end + 2;
			} else i++;
		}
	}
	code(Infinity);
	return templates.filter((text) => /[^{}]+\{[^{}]*:[^{}]*\}/.test(text)).join('\n');
}
