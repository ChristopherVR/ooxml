// Static check for formulas whose result depends on dynamic-array evaluation, so a writer can
// store them as dynamic arrays and a pre-dynamic-array reader does not reduce them by implicit
// intersection.
import { FormulaError, type FormulaAst } from './ast.js';
import { getFunction } from './functions/registry.js';
import { paramKind } from './functions/types.js';
import { parseFormula } from './parser.js';

/** Functions that evaluate their arguments as arrays even in pre-dynamic-array Excel. */
const NATIVE_ARRAY = new Set(['SUMPRODUCT', 'MMULT', 'TRANSPOSE', 'LOOKUP', 'INDEX', 'MATCH']);

interface Scan {
	/** The node can produce more than one value. */
	multi: boolean;
	/** Some sub-expression reduces a multi-valued operand where legacy Excel would intersect. */
	needs: boolean;
}

function scan(node: FormulaAst, native: boolean): Scan {
	switch (node.type) {
		case 'ref':
			return { multi: !!node.spill || (!!node.ref && node.ref.kind !== 'cell'), needs: false };
		case 'array':
			return { multi: node.rows.length > 1 || (node.rows[0]?.length ?? 0) > 1, needs: false };
		case 'structured':
			return { multi: !node.ref.specials.includes('#This Row'), needs: false };
		case 'unary':
		case 'percent': {
			const inner = scan(node.operand, native);
			if (node.type === 'unary' && node.op === '@') return { multi: false, needs: inner.needs };
			return { multi: inner.multi, needs: inner.needs || (inner.multi && !native) };
		}
		case 'binary': {
			const l = scan(node.left, native);
			const r = scan(node.right, native);
			if (node.op === ':' || node.op === ' ' || node.op === ',')
				return { multi: true, needs: l.needs || r.needs };
			const multi = l.multi || r.multi;
			return { multi, needs: l.needs || r.needs || (multi && !native) };
		}
		case 'call': {
			const spec = getFunction(node.name);
			const inner = native || NATIVE_ARRAY.has(node.name);
			let needs = false;
			let multi = false;
			node.args.forEach((arg, i) => {
				const s = scan(arg, inner);
				needs ||= s.needs;
				if (spec && !spec.lazy && paramKind(spec, i) === 'value' && s.multi) {
					multi = true;
					if (!inner) needs = true;
				}
			});
			return { multi, needs };
		}
		case 'invoke':
			return { multi: true, needs: true };
		default:
			return { multi: false, needs: false };
	}
}

/**
 * Whether a formula needs dynamic-array evaluation to keep its meaning: it returns several
 * values, or applies an operator or a single-value parameter to a range or array outside the
 * functions that take arrays natively (`SUM(A1:A3*B1:B3)`). Unparsable formulas report false.
 */
export function needsArrayEvaluation(formula: string): boolean {
	let ast: FormulaAst;
	try {
		ast = parseFormula(formula);
	} catch (e) {
		if (e instanceof FormulaError) return false;
		throw e;
	}
	const result = scan(ast, false);
	return result.needs || result.multi;
}
