import { translateFormula } from '../formula/transform.js';
import { fromStoredSyntax, toStoredSyntax } from './formula-storage.js';

/**
 * Functions added after Excel 2007. Files store them with an `_xlfn.` prefix (FILTER and SORT
 * with `_xlfn._xlws.`); the model keeps the plain name and the writer adds the prefix back.
 */
const FUTURE_FUNCTIONS = new Set(
	(
		'AGGREGATE BETA.DIST BETA.INV BINOM.DIST BINOM.INV CEILING.PRECISE CHISQ.DIST CHISQ.DIST.RT ' +
		'CHISQ.INV CHISQ.INV.RT CHISQ.TEST CONFIDENCE.NORM CONFIDENCE.T COVARIANCE.P COVARIANCE.S ' +
		'ERF.PRECISE ERFC.PRECISE EXPON.DIST F.DIST F.DIST.RT F.INV F.INV.RT F.TEST FLOOR.PRECISE ' +
		'GAMMA.DIST GAMMA.INV GAMMALN.PRECISE HYPGEOM.DIST LOGNORM.DIST LOGNORM.INV MODE.MULT MODE.SNGL ' +
		'NEGBINOM.DIST NETWORKDAYS.INTL NORM.DIST NORM.INV NORM.S.DIST NORM.S.INV PERCENTILE.EXC ' +
		'PERCENTILE.INC PERCENTRANK.EXC PERCENTRANK.INC POISSON.DIST QUARTILE.EXC QUARTILE.INC RANK.AVG ' +
		'RANK.EQ STDEV.P STDEV.S T.DIST T.DIST.2T T.DIST.RT T.INV T.INV.2T T.TEST VAR.P VAR.S ' +
		'WEIBULL.DIST WORKDAY.INTL ACOT ACOTH ARABIC BASE BINOM.DIST.RANGE BITAND BITLSHIFT BITOR ' +
		'BITRSHIFT BITXOR CEILING.MATH COMBINA COT COTH CSC CSCH DAYS DECIMAL ENCODEURL FILTERXML ' +
		'FLOOR.MATH FORMULATEXT GAMMA GAUSS IFNA IMCOSH IMCOT IMCSC IMCSCH IMSEC IMSECH IMSINH IMTAN ' +
		'ISFORMULA ISOWEEKNUM MUNIT NUMBERVALUE PDURATION PERMUTATIONA PHI RRI SEC SECH SHEET SHEETS ' +
		'SKEW.P UNICHAR UNICODE WEBSERVICE XOR CONCAT FORECAST.ETS FORECAST.ETS.CONFINT ' +
		'FORECAST.ETS.SEASONALITY FORECAST.ETS.STAT FORECAST.LINEAR IFS MAXIFS MINIFS SWITCH TEXTJOIN ' +
		'XLOOKUP XMATCH FILTER SORT SORTBY UNIQUE SEQUENCE RANDARRAY LET LAMBDA TEXTBEFORE TEXTAFTER ' +
		'TEXTSPLIT VSTACK HSTACK TAKE DROP CHOOSEROWS CHOOSECOLS TOCOL TOROW WRAPROWS WRAPCOLS EXPAND ' +
		'ARRAYTOTEXT VALUETOTEXT ISOMITTED BYROW BYCOL MAP REDUCE SCAN MAKEARRAY IMAGE GROUPBY PIVOTBY ' +
		'PERCENTOF REGEXTEST REGEXEXTRACT REGEXREPLACE TRIMRANGE ANCHORARRAY SINGLE STOCKHISTORY ' +
		'FIELDVALUE'
	).split(' '),
);
const WORKSHEET_PREFIXED = new Set(['FILTER', 'SORT']);

type TokenVisitor = (token: string, next: string, out: string[]) => boolean;

/**
 * Walks a formula outside string literals, quoted sheet names and structured-reference
 * brackets, offering every identifier-like token to `visit`; tokens it declines are copied.
 */
function rewrite(formula: string, visit: TokenVisitor): string {
	const out: string[] = [];
	let i = 0;
	while (i < formula.length) {
		const ch = formula[i] ?? '';
		if (ch === '"' || ch === "'") {
			let j = i + 1;
			while (j < formula.length) {
				if (formula[j] === ch) {
					if (formula[j + 1] === ch) j += 2;
					else break;
				} else j++;
			}
			out.push(formula.slice(i, j + 1));
			i = j + 1;
			continue;
		}
		if (ch === '[') {
			let depth = 0;
			let j = i;
			for (; j < formula.length; j++) {
				if (formula[j] === '[') depth++;
				else if (formula[j] === ']' && --depth === 0) break;
			}
			out.push(formula.slice(i, j + 1));
			i = j + 1;
			continue;
		}
		if (/[A-Za-z0-9_$\\.]/.test(ch)) {
			let j = i;
			while (j < formula.length && /[A-Za-z0-9_$\\.?]/.test(formula[j] ?? '')) j++;
			const token = formula.slice(i, j);
			if (!visit(token, formula.slice(j), out)) out.push(token);
			i = j;
			continue;
		}
		out.push(ch);
		i++;
	}
	return out.join('');
}

/** Removes `_xlfn.` / `_xlws.` function prefixes (and turns `ANCHORARRAY(A1)` back into `A1#`). */
export function stripFuturePrefixes(formula: string): string {
	if (!formula.includes('_xl')) return formula;
	const plain = rewrite(formula, (token, next, out) => {
		if (!next.startsWith('(')) return false;
		const stripped = token.replace(/^(_xlfn\.)?(_xlws\.)?/i, '');
		if (stripped === token) return false;
		out.push(stripped);
		return true;
	});
	return fromStoredSyntax(plain);
}

/**
 * Adds `_xlfn.` prefixes to post-2007 functions, as Excel expects them in the file, writing the
 * spill operator `A1#` as `ANCHORARRAY(A1)` and `@x` as `SINGLE(x)`.
 */
export function addFuturePrefixes(formula: string): string {
	return rewrite(toStoredSyntax(formula), (token, next, out) => {
		if (!next.startsWith('(')) return false;
		const upper = token.toUpperCase();
		if (!FUTURE_FUNCTIONS.has(upper)) return false;
		out.push(WORKSHEET_PREFIXED.has(upper) ? `_xlfn._xlws.${token}` : `_xlfn.${token}`);
		return true;
	});
}

/**
 * Moves the relative references of a formula by `dRow` rows and `dCol` columns, as Excel does
 * when a shared formula applies to another cell. Absolute (`$`) parts stay put; references
 * pushed off the grid become `#REF!`. Kept for compatibility: an alias of the formula module's
 * {@link translateFormula}.
 */
export function translateReferences(formula: string, dRow: number, dCol: number): string {
	return translateFormula(formula, dRow, dCol);
}
