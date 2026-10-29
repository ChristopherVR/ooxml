import { themeInputs } from './formula-theme';
import { numericFormulaFunctions } from './formula-numeric-functions';
import { isTheText, textNumberFunctions, textResultFunctions } from './formula-text-functions';
import { FRAME_INPUTS, literalIndirectTarget } from './formula-indirect';
import { mathFormulaFunctions } from './formula-math-functions';
import { formulaFailure, formulaLimit, parseVisioFormula } from './formula';
import type { VisioFormulaAst, VisioFormulaLimits, VisioFormulaReference } from './formula';

export const supported = new Set([
	'GUARD',
	'IF',
	'ABS',
	'MIN',
	'MAX',
	'BOUND',
	'SQRT',
	'SIN',
	'COS',
	'TAN',
	'ATAN2',
	'SIGN',
	'PI',
	'PAGENUMBER',
	'PAGECOUNT',
	'PNT',
	'PNTX',
	'PNTY',
	'ISERROR',
	'IFERROR',
	'INDIRECT',
	'DATE',
	'ANGLETOLOC',
	'LOCTOPAR',
	'SHEETREF',
	...textNumberFunctions,
	...textResultFunctions,
	...numericFormulaFunctions,
	...mathFormulaFunctions,
]);
// Known pure functions may be independent of an edit even when evaluation is unsupported.
const staticUnsupported = new Set([
	// FONT reads the font table by name, not a hidden ShapeSheet reference.
	// Keep evaluation unsupported: identifiers depend on the document/system.
	// https://learn.microsoft.com/en-us/office/client-developer/visio/font-function
	'FONT',
	// Page-name lookup has explicit arguments but returns a string, outside numeric evaluation.
	'PAGENAME',
	// Text-field functions read document properties or the clock, never ShapeSheet cells.
	// https://learn.microsoft.com/en-us/office/client-developer/visio/now-function
	'NOW',
	'DOCCREATION',
	'DOCLASTSAVE',
	'DOCLASTEDIT',
	'DOCLASTPRINT',
	'TITLE',
	'SUBJECT',
	'CREATOR',
	'KEYWORDS',
	'DESCRIPTION',
	'CATEGORY',
	'COMPANY',
	'MANAGER',
	// POLYLINE encodes explicitly supplied geometry data; affected expressions
	// still require the geometry evaluator and are rejected by numeric recalculation.
	// https://learn.microsoft.com/en-us/office/client-developer/visio/polyline-function
	'POLYLINE',
	// NURBS likewise carries explicit control points (curved connectors write literal ones).
	// https://learn.microsoft.com/en-us/office/client-developer/visio/nurbs-function
	'NURBS',
	'MOD',
	'LOG',
	'AND',
	'OR',
	'NOT',
	'RGB',
	'HSL',
	'THEMEGUARD',
	'THEMEVAL',
	'THEME',
	'ISTHEMED',
	'BITAND',
	'BITOR',
	'BITXOR',
	'TINT',
	'MSOTINT',
	'SHADE',
	'LUMDIFF',
	'MODULUS',
	// Reads an argument of the master the shape was dropped from; pure, evaluated in master scope.
	'ARG',
]);
// These functions can introduce references that cannot be identified from argument ASTs.
const dynamic = new Set([
	'EVALCELL',
	'EVALTEXT',
	'DEPENDSON',
	'SETATREF',
	'SETATREFEXPR',
	'SETATREFEVAL',
	'GETREF',
	'REF',
	'SHAPETEXT',
	'CALLOUTTARGET',
	'CONTAINERSHEETREF',
	'MEMBERSHEETREF',
	'PARENT',
	'PAR',
]);
/** A function the analyser classifies by name, whether or not it can evaluate it. */
export const isClassifiedVisioFunction = (name: string): boolean =>
	supported.has(name) || staticUnsupported.has(name) || dynamic.has(name);
export interface VisioFormulaAnalysis {
	references: VisioFormulaReference[];
	unsupportedFunctions: string[];
	dynamic: boolean;
	/**
	 * The formula can read a cell its text does not name (EVALCELL, a computed INDIRECT, a cell of a
	 * returned sheet, an unknown function). A merely dynamic formula still names what it reads.
	 */
	opaque: boolean;
	guarded: boolean;
}
/** Functions whose reads cannot be recovered from their arguments. */
const opaqueFunctions = new Set([
	'EVALCELL',
	'EVALTEXT',
	'GETREF',
	'REF',
	'_SHEETCELL',
	// Structure lookups: the result changes when shapes are created, moved or reparented.
	'CONTAINERSHEETREF',
	'MEMBERSHEETREF',
	'CALLOUTTARGET',
	'PARENT',
	'PAR',
]);
export function traverse(
	ast: VisioFormulaAst,
	limits: VisioFormulaLimits,
	visit: (node: VisioFormulaAst) => void,
): void {
	const maxNodes = formulaLimit(limits.maxNodes, 1024),
		maxDepth = formulaLimit(limits.maxDepth, 64);
	const queue: [VisioFormulaAst, number][] = [[ast, 0]];
	let count = 0;
	while (queue.length) {
		const [node, depth] = queue.pop()!;
		if (++count > maxNodes || depth > maxDepth)
			return formulaFailure('limit', 'Formula analysis limit');
		visit(node);
		if (node.kind === 'unary') queue.push([node.operand, depth + 1]);
		if (node.kind === 'binary') queue.push([node.left, depth + 1], [node.right, depth + 1]);
		if (node.kind === 'call') for (const arg of node.args) queue.push([arg, depth + 1]);
	}
}
export function analyzeVisioFormula(
	source: string | VisioFormulaAst,
	limits: VisioFormulaLimits = {},
): VisioFormulaAnalysis {
	const ast = typeof source === 'string' ? parseVisioFormula(source, limits) : source;
	const references = new Map<string, VisioFormulaReference>(),
		unsupportedFunctions = new Set<string>();
	let unsafe = false,
		opaque = false,
		guarded = false;
	traverse(ast, limits, (node) => {
		if (node.kind === 'reference')
			references.set(
				`${node.reference.shapeId ?? ''}!${node.reference.cell.toLowerCase()}`,
				node.reference,
			);
		if (node.kind === 'call') {
			const inputs =
				node.name === 'ISTHEMED'
					? ['ColorSchemeIndex']
					: ['THEMEVAL', 'THEME'].includes(node.name)
						? themeInputs
						: [];
			for (const cell of inputs) references.set(`!${cell.toLowerCase()}`, { cell });
			for (const cell of FRAME_INPUTS[node.name] ?? [])
				references.set(`!${cell.toLowerCase()}`, { cell });
			if (node.name === 'INDIRECT') {
				// A literal address is an ordinary static reference when the caller opts in; a computed one is
				// never provable, and by default every INDIRECT stays dynamic.
				const target = literalIndirectTarget(node);
				if (target && limits.literalIndirect)
					references.set(`${target.shapeId ?? ''}!${target.cell.toLowerCase()}`, target);
				else {
					unsafe = true;
					opaque = true;
				}
			}
			if (!supported.has(node.name)) unsupportedFunctions.add(node.name);
			const shapeText =
				node.name === 'SHAPETEXT' &&
				limits.shapeTextReads === true &&
				node.args.length === 1 &&
				isTheText(node.args[0]!);
			if (
				(dynamic.has(node.name) && !shapeText) ||
				(!supported.has(node.name) && !staticUnsupported.has(node.name))
			)
				unsafe = true;
			if (
				opaqueFunctions.has(node.name) ||
				(node.name === 'SHAPETEXT' && !shapeText) ||
				(!supported.has(node.name) &&
					!staticUnsupported.has(node.name) &&
					!dynamic.has(node.name) &&
					node.name !== '_SHEETCALL')
			)
				opaque = true;
			if (node.name === 'GUARD') guarded = true;
		}
	});
	return {
		references: [...references.values()],
		unsupportedFunctions: [...unsupportedFunctions],
		dynamic: unsafe,
		opaque,
		guarded,
	};
}
