import { describe, expect, it } from 'vitest';
import { FormulaError, type FormulaAst } from './ast.js';
import { parseFormula } from './parser.js';

/** A compact s-expression of the tree, for readable precedence assertions. */
function show(node: FormulaAst): string {
	switch (node.type) {
		case 'number':
			return String(node.value);
		case 'string':
			return JSON.stringify(node.value);
		case 'boolean':
			return node.value ? 'TRUE' : 'FALSE';
		case 'error':
			return node.code;
		case 'ref':
			return `${node.prefix ? `${node.prefix.sheet}!` : ''}${node.ref ? `R${node.ref.start.row}C${node.ref.start.col}` : '#REF!'}${node.ref && node.ref.kind !== 'cell' ? `:R${node.ref.end.row}C${node.ref.end.col}` : ''}${node.spill ? '#' : ''}`;
		case 'name':
			return `name(${node.name})`;
		case 'structured':
			return `table(${node.ref.table ?? ''}|${node.ref.specials.join('+')}|${node.ref.column ?? ''}${node.ref.column2 ? `:${node.ref.column2}` : ''})`;
		case 'array':
			return `{${node.rows.map((r) => r.map((v) => (typeof v === 'object' ? v.error : String(v))).join(',')).join(';')}}`;
		case 'unary':
			return `(${node.op} ${show(node.operand)})`;
		case 'percent':
			return `(% ${show(node.operand)})`;
		case 'binary':
			return `(${node.op === ' ' ? '_' : node.op} ${show(node.left)} ${show(node.right)})`;
		case 'call':
			return `${node.name}(${node.args.map(show).join(' ')})`;
		case 'invoke':
			return `invoke(${show(node.callee)} ${node.args.map(show).join(' ')})`;
		case 'missing':
			return '_';
	}
}

const p = (formula: string): string => show(parseFormula(formula));

describe('parseFormula precedence', () => {
	it.each([
		['1+2*3', '(+ 1 (* 2 3))'],
		['1-2-3', '(- (- 1 2) 3)'],
		['2^3^2', '(^ (^ 2 3) 2)'],
		['-2^2', '(^ (- 2) 2)'],
		['2^-2', '(^ 2 (- 2))'],
		['-5%', '(- (% 5))'],
		['50%*2', '(* (% 50) 2)'],
		['1+2&"a"', '(& (+ 1 2) "a")'],
		['"a"&1=2', '(= (& "a" 1) 2)'],
		['1<2=TRUE', '(= (< 1 2) TRUE)'],
		['A1:B2 B1:C3', '(_ R0C0:R1C1 R0C1:R2C2)'],
		['-A1:A2', '(- R0C0:R1C0)'],
		['A1:INDEX(B:B,2)', '(: R0C0 INDEX(R0C1:R1048575C1 2))'],
		['SUM((A1,B2))', 'SUM((, R0C0 R1C1))'],
		['(1)', '1'],
		['@A1:A3', '(@ R0C0:R2C0)'],
		['+A1', '(+ R0C0)'],
	])('%s', (formula, expected) => {
		expect(p(formula)).toBe(expected);
	});
});

describe('parseFormula operands', () => {
	it('parses function calls with missing arguments', () => {
		expect(p('IF(A1,,)')).toBe('IF(R0C0 _ _)');
		expect(p('PI()')).toBe('PI()');
		expect(p('SUM(1, 2 , 3)')).toBe('SUM(1 2 3)');
	});

	it('normalizes stored function prefixes but keeps the raw name', () => {
		const ast = parseFormula('_xlfn._xlws.SORT(A1:A3)');
		expect(ast).toMatchObject({ type: 'call', name: 'SORT', rawName: '_xlfn._xlws.SORT' });
		expect(p('sum(1)')).toBe('SUM(1)');
	});

	it('parses array constants with negatives, text, logicals and errors', () => {
		expect(p('{1,-2;"a",TRUE}')).toBe('{1,-2;a,true}');
		expect(p('{#N/A,1}')).toBe('{#N/A,1}');
		expect(() => parseFormula('{1,2;3}')).toThrow(FormulaError);
		expect(() => parseFormula('{A1}')).toThrow(FormulaError);
	});

	it('parses sheet-qualified references, names and #REF!', () => {
		expect(p("'My Sheet'!A1+Sheet2!B2")).toBe('(+ My Sheet!R0C0 Sheet2!R1C1)');
		expect(p('Sheet1!Rate')).toBe('name(Rate)');
		expect(p('#REF!+1')).toBe('(+ #REF! 1)');
	});

	it('parses structured references', () => {
		expect(p('Sales[Amount]')).toBe('table(Sales||Amount)');
		expect(p('Sales[[#This Row],[Amount]]')).toBe('table(Sales|#This Row|Amount)');
		expect(p('Sales[@Amount]')).toBe('table(Sales|#This Row|Amount)');
		expect(p('[@[Unit Price]]')).toBe('table(|#This Row|Unit Price)');
		expect(p('Sales[#All]')).toBe('table(Sales|#All|)');
		expect(p('Sales[[#Headers],[#Data],[Q1]:[Q4]]')).toBe('table(Sales|#Headers+#Data|Q1:Q4)');
		expect(p('Sales[]')).toBe('table(Sales||)');
	});

	it('parses LAMBDA invocation', () => {
		expect(p('LAMBDA(x,x+1)(2)')).toBe('invoke(LAMBDA(name(x) (+ name(x) 1)) 2)');
	});

	it('parses spill references', () => {
		expect(p('SUM(A1#)')).toBe('SUM(R0C0#)');
	});

	it('ignores insignificant whitespace and newlines', () => {
		expect(p(' 1 +\n 2 ')).toBe('(+ 1 2)');
		expect(p('IF( A1 > 0 , 1 , 2 )')).toBe('IF((> R0C0 0) 1 2)');
	});

	it.each(['', '1+', '(1', 'SUM(1', '1 2', ')', 'SUM(1,,', '*2'])('rejects %j', (formula) => {
		expect(() => parseFormula(formula)).toThrow(FormulaError);
	});
});
