import { formulaFailure } from './formula';
import type { VisioFormulaAst, VisioFormulaReference, VisioFormulaValue } from './formula';
import { formatNumberPicture, strSameEx } from './formula-extra-functions';

/**
 * ShapeSheet string functions that work on text written in the formula itself. A cell's own text
 * (`TheText`, `Prop.Name`) is not known to the numeric evaluator, so a formula that reads one is
 * still refused. Behaviour at the edges (an out-of-range `MID` start, a missing `FIND`) was
 * recorded from Visio 16 with `scripts/record-visio-formula-functions.ps1`.
 * https://learn.microsoft.com/en-us/office/client-developer/visio/mid-function
 */

/** Functions that return text; they may only appear as arguments of a numeric function. */
export const textResultFunctions: ReadonlySet<string> = new Set([
	'MID',
	'LEFT',
	'RIGHT',
	'UPPER',
	'LOWER',
	'TRIM',
	'SUBSTITUTE',
	'REPLACE',
	'CHAR',
	'INDEX',
	'FORMAT',
	'SHAPETEXT',
]);
/** Functions that read text and return a number. */
export const textNumberFunctions: ReadonlySet<string> = new Set([
	'LEN',
	'FIND',
	'STRSAME',
	'STRSAMEEX',
	'LOOKUP',
]);

/** The list separator Visio uses between the items of `LOOKUP` and `INDEX` lists. */
const LIST_SEPARATOR = ';';

export interface VisioTextContext {
	/** Evaluate a numeric argument (counted against the evaluation limits). */
	number(node: VisioFormulaAst): VisioFormulaValue;
	/** The plain text of a `TheText` reference, or undefined when it cannot be known. */
	textOf?(reference: VisioFormulaReference): string | undefined;
	/** Count one evaluation step. */
	tick(): void;
}

const unsupported = (message: string): never => formulaFailure('unsupported', message);

function integer(context: VisioTextContext, node: VisioFormulaAst, name: string): number {
	const value = context.number(node);
	if (value.unit !== 'scalar') return formulaFailure('unit', `${name} requires a plain number`);
	return Math.trunc(value.value);
}

/** A whole-number argument that must not have a fractional part (rounding is unrecorded). */
function exactInteger(context: VisioTextContext, node: VisioFormulaAst, name: string): number {
	const value = context.number(node);
	if (value.unit !== 'scalar') return formulaFailure('unit', `${name} requires a plain number`);
	if (!Number.isInteger(value.value)) return unsupported(`${name} with a fractional number`);
	return value.value;
}

/** True for `TheText` and `Sheet.N!TheText`. */
export const isTheText = (node: VisioFormulaAst): boolean =>
	node.kind === 'reference' && node.reference.cell.toLowerCase() === 'thetext';

/** The text of an argument: a literal, a text function, `&`, or a plain number written out. */
export function evaluateFormulaText(
	node: VisioFormulaAst,
	context: VisioTextContext,
	depth = 0,
): string {
	context.tick();
	if (depth > 64) return formulaFailure('limit', 'Formula evaluation limit');
	const text = (child: VisioFormulaAst | undefined, name: string) => {
		if (!child) return formulaFailure('arity', `Invalid ${name} argument count`);
		return evaluateFormulaText(child, context, depth + 1);
	};
	if (node.kind === 'string') return node.value;
	if (node.kind === 'binary' && node.operator === '&')
		return text(node.left, '&') + text(node.right, '&');
	if (node.kind === 'reference' && isTheText(node)) {
		const read = context.textOf?.(node.reference);
		return read === undefined ? unsupported('Text read from a cell') : read;
	}
	if (node.kind !== 'call') {
		// A number is written the way Visio writes it in a joined string: plain, no unit.
		let value: VisioFormulaValue;
		try {
			value = context.number(node);
		} catch (error) {
			if (node.kind === 'reference') return unsupported('Text read from a cell');
			throw error;
		}
		if (value.unit !== 'scalar') return unsupported('Text from a number with a unit');
		return String(value.value);
	}
	const { name, args } = node;
	const arity = (min: number, max = min) => {
		if (args.length < min || args.length > max)
			return formulaFailure('arity', `Invalid ${name} argument count`);
	};
	if (!textResultFunctions.has(name)) {
		// A numeric call (a number joined into a string, INT(...) and so on).
		const value = context.number(node);
		if (value.unit !== 'scalar') return unsupported('Text from a number with a unit');
		return String(value.value);
	}
	if (name === 'SHAPETEXT') {
		// SHAPETEXT(TheText) or SHAPETEXT(Sheet.N!TheText): the plain text of a shape.
		arity(1);
		const target = args[0]!;
		if (!isTheText(target)) return unsupported('SHAPETEXT of something other than TheText');
		return text(target, name);
	}
	if (name === 'MID') {
		arity(3);
		const source = text(args[0], name),
			start = integer(context, args[1]!, name),
			length = integer(context, args[2]!, name);
		if (start < 1) return formulaFailure('value', 'MID start must be at least 1');
		// A negative length takes the rest of the string, as Visio does.
		return length < 0 ? source.slice(start - 1) : source.slice(start - 1, start - 1 + length);
	}
	if (name === 'LEFT' || name === 'RIGHT') {
		arity(1, 2);
		const source = text(args[0], name);
		const count = args[1] ? exactInteger(context, args[1], name) : 1;
		if (count < 0) {
			// Recorded for LEFT only: a negative count returns the whole string.
			if (name === 'RIGHT') return unsupported('RIGHT with a negative count');
			return source;
		}
		if (count === 0) return '';
		return name === 'LEFT' ? source.slice(0, count) : source.slice(-count);
	}
	if (name === 'UPPER' || name === 'LOWER') {
		arity(1);
		const source = text(args[0], name);
		return name === 'UPPER' ? source.toUpperCase() : source.toLowerCase();
	}
	if (name === 'TRIM') {
		arity(1);
		// Leading and trailing spaces go, and runs of spaces inside collapse to one.
		return text(args[0], name).replace(/ {2,}/g, ' ').replace(/^ | $/g, '');
	}
	if (name === 'SUBSTITUTE') {
		arity(3, 4);
		const source = text(args[0], name),
			find = text(args[1], name),
			replacement = text(args[2], name);
		if (find === '') return source;
		if (!args[3]) return source.split(find).join(replacement);
		const instance = exactInteger(context, args[3], name);
		if (instance < 1) return unsupported('SUBSTITUTE with an instance below 1');
		let at = -1;
		for (let seen = 0; seen < instance; seen++) {
			at = source.indexOf(find, at + 1);
			if (at < 0) return source;
		}
		return source.slice(0, at) + replacement + source.slice(at + find.length);
	}
	if (name === 'REPLACE') {
		arity(4);
		const source = text(args[0], name),
			start = integer(context, args[1]!, name),
			count = integer(context, args[2]!, name),
			insert = text(args[3], name);
		if (start < 1) return formulaFailure('value', 'REPLACE start must be at least 1');
		if (count < 0) return unsupported('REPLACE with a negative count');
		return source.slice(0, start - 1) + insert + source.slice(start - 1 + count);
	}
	if (name === 'CHAR') {
		arity(1);
		const code = integer(context, args[0]!, name);
		if (code < 1 || code > 255) return formulaFailure('value', 'CHAR requires 1 to 255');
		// 128 to 159 are Windows-1252 characters that Latin-1 does not hold.
		if (code >= 128 && code <= 159) return unsupported('CHAR in the Windows-1252 range');
		return String.fromCharCode(code);
	}
	if (name === 'FORMAT') {
		// FORMAT(number, picture): numeric pictures only (formula-extra-functions.ts).
		arity(2);
		return formatNumberPicture(context.number(args[0]!), text(args[1], name));
	}
	// INDEX(i, list): the zero-based item, or an empty string when there is none.
	arity(2);
	const index = integer(context, args[0]!, name);
	const items = text(args[1], name).split(LIST_SEPARATOR);
	return index >= 0 && index < items.length ? items[index]! : '';
}

/** `LEN`, `FIND`, `STRSAME` and `LOOKUP`: numeric results of text arguments. */
export function evaluateTextNumberFunction(
	name: string,
	args: readonly VisioFormulaAst[],
	context: VisioTextContext,
): VisioFormulaValue {
	const arity = (min: number, max = min) => {
		if (args.length < min || args.length > max)
			return formulaFailure('arity', `Invalid ${name} argument count`);
	};
	const text = (index: number) => evaluateFormulaText(args[index]!, context, 1);
	const scalar = (value: number): VisioFormulaValue => ({ value, unit: 'scalar' });
	if (name === 'LEN') {
		arity(1);
		return scalar(text(0).length);
	}
	if (name === 'FIND') {
		// FIND(find, within[, start]): 1-based, case-sensitive, #VALUE! when it is not there.
		arity(2, 3);
		const find = text(0),
			within = text(1);
		const start = args[2] ? integer(context, args[2], name) : 1;
		if (start < 1 || start > within.length)
			return formulaFailure('value', 'FIND start is outside the text');
		const at = within.indexOf(find, start - 1);
		if (at < 0) return formulaFailure('value', 'FIND did not find the text');
		return scalar(at + 1);
	}
	if (name === 'STRSAME') {
		// A third argument other than 0 ignores case.
		arity(2, 3);
		const a = text(0),
			b = text(1);
		const ignoreCase = args[2] ? context.number(args[2]).value !== 0 : false;
		return scalar(ignoreCase ? Number(a.toLowerCase() === b.toLowerCase()) : Number(a === b));
	}
	if (name === 'STRSAMEEX') {
		// STRSAMEEX(a, b, flags): flags are Windows NORM_* bits; a locale argument is unsupported.
		if (args.length === 4) return formulaFailure('unsupported', 'STRSAMEEX with a locale');
		arity(3);
		return scalar(Number(strSameEx(text(0), text(1), context.number(args[2]!).value)));
	}
	// LOOKUP(text, list): the zero-based place of the first item that matches without regard to
	// case, or -1.
	arity(2);
	const wanted = text(0).toLowerCase();
	return scalar(
		text(1)
			.split(LIST_SEPARATOR)
			.findIndex((item) => item.toLowerCase() === wanted),
	);
}
