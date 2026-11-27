/**
 * A display paragraph (`m:oMathPara`) that holds several equations (`m:oMath`, one per line as Word
 * shows them) is edited one equation at a time: each one is read, rebuilt from its linear text and
 * written back into the paragraph's markup in place, so the paragraph's properties and every other
 * equation stay byte for byte as they were. Why an equation stays read-only is told apart, so the
 * editors can say what keeps it as it is.
 */
import { controlFormat, hasTrackedChanges, sharedRunFormat } from './equation-canonical';
import { editableEquation, equationMath, type EquationInputFormat } from './linear-format';
import { parseOmml } from './omml-dom';

/** An element's span in markup: `start` at its `<`, `end` just past its closing `>`. */
interface ElementSpan {
	name: string;
	start: number;
	end: number;
}

const NAME = /^<\/?([^\s/>]+)/;

/**
 * The spans of the root element's direct child elements, in order. Comments, CDATA, processing
 * instructions and quoted attribute values are skipped; undefined for markup it cannot follow.
 */
function childSpans(xml: string): ElementSpan[] | undefined {
	const spans: ElementSpan[] = [];
	let depth = 0;
	let open: ElementSpan | undefined;
	let at = xml.indexOf('<');
	while (at >= 0 && at < xml.length) {
		if (xml.startsWith('<!--', at)) at = after(xml, '-->', at);
		else if (xml.startsWith('<![CDATA[', at)) at = after(xml, ']]>', at);
		else if (xml.startsWith('<?', at) || xml.startsWith('<!', at)) at = after(xml, '>', at);
		else {
			const end = tagEnd(xml, at);
			if (end < 0) return undefined;
			const tag = xml.slice(at, end);
			const name = NAME.exec(tag)?.[1] ?? '';
			if (tag.startsWith('</')) {
				depth--;
				if (depth === 1 && open) {
					open.end = end;
					spans.push(open);
					open = undefined;
				}
			} else if (tag.endsWith('/>')) {
				if (depth === 1) spans.push({ name, start: at, end });
			} else {
				if (depth === 1) open = { name, start: at, end: -1 };
				depth++;
			}
			at = end;
		}
		if (at < 0) return undefined;
		at = xml.indexOf('<', at);
	}
	return depth === 0 ? spans : undefined;
}

const after = (xml: string, close: string, from: number) => {
	const found = xml.indexOf(close, from);
	return found < 0 ? -1 : found + close.length;
};

/** The index just past the `>` that ends the tag at `from`, outside quoted attribute values. */
function tagEnd(xml: string, from: number): number {
	let quote = '';
	for (let at = from + 1; at < xml.length; at++) {
		const char = xml[at];
		if (quote) {
			if (char === quote) quote = '';
		} else if (char === '"' || char === "'") quote = char;
		else if (char === '>') return at + 1;
	}
	return -1;
}

const localName = (name: string) => name.slice(name.indexOf(':') + 1);

/** The spans of a display paragraph's equations; undefined when the markup is not one. */
function partSpans(omml: string): ElementSpan[] | undefined {
	const root = NAME.exec(omml.slice(omml.indexOf('<')))?.[1];
	if (!root || localName(root) !== 'oMathPara') return undefined;
	return childSpans(omml)?.filter((span) => localName(span.name) === 'oMath');
}

/**
 * How many equations an equation's markup holds: the `m:oMath` lines of a display paragraph, or 1
 * for a single equation (or markup that cannot be read, which stays one read-only equation).
 */
export function equationParts(omml: string): number {
	return Math.max(1, partSpans(omml)?.length ?? 1);
}

/**
 * The paragraph `omml` with its equation at `part` replaced by `math` (an `m:oMath` element, as the
 * linear editor writes an inline equation). The namespace declarations `math` repeats from the
 * paragraph are dropped; everything else in the paragraph keeps its markup exactly. Undefined when
 * there is no such equation or `math` is not one.
 */
export function replaceEquationPart(omml: string, part: number, math: string): string | undefined {
	const span = partSpans(omml)?.[part];
	const head = NAME.exec(math.trimStart())?.[1];
	if (!span || !head || localName(head) !== 'oMath') return undefined;
	const rootEnd = tagEnd(omml, omml.indexOf('<'));
	const rootTag = rootEnd < 0 ? '' : omml.slice(0, rootEnd);
	const trimmed = math.trim();
	const end = tagEnd(trimmed, 0);
	if (end < 0) return undefined;
	const tag = trimmed
		.slice(0, end)
		.replace(/\s+xmlns:([\w.-]+)=("[^"]*"|'[^']*')/g, (declaration, prefix: string, value) =>
			rootTag.includes(`xmlns:${prefix}=${value}`) ? '' : declaration,
		);
	return omml.slice(0, span.start) + tag + trimmed.slice(end) + omml.slice(span.end);
}

/** Why an equation stays read-only in the linear editor. */
export type EquationReadOnlyReason = 'tracked' | 'formatting' | 'content';

/**
 * Why the equation (or its equation at `part`) cannot be edited as linear text, or undefined when
 * it can: `tracked` for tracked changes inside it, `formatting` for runs formatted differently (or
 * control characters with more than the math font), `content` for anything else its notation has
 * no spelling for.
 */
export function equationReadOnlyReason(
	omml: string,
	format: EquationInputFormat,
	part?: number,
): EquationReadOnlyReason | undefined {
	if (editableEquation(omml, format, part)) return undefined;
	try {
		const math = equationMath(parseOmml(omml), part)?.math;
		if (!math) return 'content';
		if (hasTrackedChanges(math)) return 'tracked';
		if (sharedRunFormat(math) === 'mixed' || controlFormat(math) === 'other') return 'formatting';
	} catch {
		// Unreadable markup has nothing more specific to say.
	}
	return 'content';
}
