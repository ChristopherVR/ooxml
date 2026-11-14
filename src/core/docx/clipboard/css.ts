/** CSS reading for Word's clipboard HTML: inline declarations, its stylesheet and its lengths. */

/** `name:value` pairs of an inline `style` (names lower-cased, last declaration wins). */
export function parseDeclarations(style: string | undefined): Map<string, string> {
	const declarations = new Map<string, string>();
	if (!style) return declarations;
	// Word wraps long attributes over lines and quotes font names; neither contains a `;`.
	for (const part of splitOutsideQuotes(style, ';')) {
		const colon = part.indexOf(':');
		if (colon <= 0) continue;
		const name = part.slice(0, colon).trim().toLowerCase();
		const value = part
			.slice(colon + 1)
			.replace(/\s+/g, ' ')
			.trim();
		if (name) declarations.set(name, value);
	}
	return declarations;
}

function splitOutsideQuotes(value: string, separator: string): string[] {
	const parts: string[] = [];
	let quote = '';
	let current = '';
	for (let index = 0; index < value.length; index++) {
		const char = value[index]!;
		if (char === '\\' && index + 1 < value.length) {
			current += char + value[++index];
			continue;
		}
		if (quote) {
			if (char === quote) quote = '';
		} else if (char === '"' || char === "'") quote = char;
		else if (char === separator) {
			parts.push(current);
			current = '';
			continue;
		}
		current += char;
	}
	parts.push(current);
	return parts;
}

/** Serialises declarations back to an inline style. */
export function serializeDeclarations(declarations: ReadonlyMap<string, string>): string {
	return [...declarations].map(([name, value]) => `${name}:${value}`).join(';');
}

/** A CSS string or identifier value without its quotes and escapes (`"%1\)"` is `%1)`). */
export function cssStringValue(value: string): string {
	const trimmed = value.trim();
	const quoted = /^(["'])(.*)\1$/s.exec(trimmed);
	const body = quoted ? quoted[2]! : trimmed;
	return body.replace(/\\([0-9a-f]{1,6}\s?|.)/gi, (_whole, escaped: string) =>
		/^[0-9a-f]+\s?$/i.test(escaped) && escaped.trim().length > 1
			? String.fromCodePoint(parseInt(escaped, 16))
			: escaped,
	);
}

const POINTS_PER_UNIT: Record<string, number> = {
	pt: 1,
	in: 72,
	cm: 72 / 2.54,
	mm: 72 / 25.4,
	pc: 12,
	px: 0.75,
};

/** A length (`.25in`, `-18.0pt`, `1.0cm`) in points, or `undefined` for other values. */
export function cssPoints(value: string | undefined): number | undefined {
	const match = /^\s*(-?(?:\d+\.?\d*|\.\d+))\s*(pt|in|cm|mm|pc|px)?\s*$/i.exec(value ?? '');
	if (!match) return undefined;
	const number = Number(match[1]);
	if (!Number.isFinite(number)) return undefined;
	if (!match[2]) return number === 0 ? 0 : undefined;
	return number * POINTS_PER_UNIT[match[2].toLowerCase()]!;
}

/** A length in whole twips. */
export function cssTwips(value: string | undefined): number | undefined {
	const points = cssPoints(value);
	return points === undefined ? undefined : Math.round(points * 20);
}

/** The first family of a `font-family` list, unquoted. */
export function firstFontFamily(value: string | undefined): string | undefined {
	const first = value ? cssStringValue(splitOutsideQuotes(value, ',')[0] ?? '') : '';
	return first || undefined;
}

export interface CssRule {
	/** The selector text (an `@list l0:level1` rule keeps its at-keyword). */
	selector: string;
	declarations: Map<string, string>;
}

/** The rules of every `<style>` block in `html`, comments removed. */
export function readStyleRules(html: string): CssRule[] {
	const rules: CssRule[] = [];
	for (const block of html.matchAll(/<style[^>]*>([\s\S]*?)<\/style>/gi)) {
		const css = block[1]!.replace(/<!--|-->/g, '').replace(/\/\*[\s\S]*?\*\//g, '');
		for (const rule of css.matchAll(/([^{}]+)\{([^}]*)\}/g)) {
			const selector = rule[1]!.replace(/\s+/g, ' ').trim();
			if (selector) rules.push({ selector, declarations: parseDeclarations(rule[2]) });
		}
	}
	return rules;
}
