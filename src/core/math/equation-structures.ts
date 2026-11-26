/**
 * The Structures galleries of the Equation tab: each structure as LaTeX whose `⬚` marks the
 * placeholders a new structure waits on. The products insert them through the shared converter
 * (`buildUpEquation`), in whichever input format the editor is set to.
 */

export type EquationStructureGroup =
	| 'fraction'
	| 'script'
	| 'radical'
	| 'integral'
	| 'largeOperator'
	| 'bracket'
	| 'function'
	| 'accent'
	| 'limitLog'
	| 'operator'
	| 'matrix';

export interface EquationStructure {
	/** Stable id within its gallery. */
	id: string;
	/** English name; the editors translate it. */
	name: string;
	latex: string;
}

export interface EquationStructureGallery {
	id: EquationStructureGroup;
	/** English label of the ribbon command. */
	name: string;
	structures: readonly EquationStructure[];
}

const P = '⬚';
const s = (id: string, name: string, latex: string): EquationStructure => ({ id, name, latex });

export const EQUATION_STRUCTURES: readonly EquationStructureGallery[] = [
	{
		id: 'fraction',
		name: 'Fraction',
		structures: [
			s('stacked', 'Stacked Fraction', `\\frac{${P}}{${P}}`),
			s('linear', 'Linear Fraction', `\\sfrac{${P}}{${P}}`),
			s('noBar', 'Stacked Fraction Without Bar', `{${P} \\atop ${P}}`),
			s('dydx', 'Differential', '\\frac{dy}{dx}'),
			s('deltas', 'Change in y over x', '\\frac{\\Delta y}{\\Delta x}'),
			s('partial', 'Partial Differential', '\\frac{\\partial y}{\\partial x}'),
			s('piOver2', 'Pi Over 2', '\\frac{\\pi}{2}'),
		],
	},
	{
		id: 'script',
		name: 'Script',
		structures: [
			s('superscript', 'Superscript', `${P}^{${P}}`),
			s('subscript', 'Subscript', `${P}_{${P}}`),
			s('subSup', 'Subscript-Superscript', `${P}_{${P}}^{${P}}`),
			s('leftSubSup', 'Left Subscript-Superscript', `{}_{${P}}^{${P}}{${P}}`),
			s('xSubY2', 'Subscript with Superscript', 'x_{y^{2}}'),
			s('eMinusIwt', 'Exponential', 'e^{-i\\omega t}'),
			s('xSquared', 'x Squared', 'x^{2}'),
		],
	},
	{
		id: 'radical',
		name: 'Radical',
		structures: [
			s('squareRoot', 'Square Root', `\\sqrt{${P}}`),
			s('withDegree', 'Radical with Degree', `\\sqrt[${P}]{${P}}`),
			s('squareRoot2', 'Square Root with Degree', `\\sqrt[2]{${P}}`),
			s('cubeRoot', 'Cube Root', `\\sqrt[3]{${P}}`),
			s('quadratic', 'Quadratic Root', '\\frac{-b\\pm\\sqrt{b^{2}-4ac}}{2a}'),
			s('hypotenuse', 'Root of a Sum of Squares', '\\sqrt{a^{2}+b^{2}}'),
		],
	},
	{
		id: 'integral',
		name: 'Integral',
		structures: [
			s('integral', 'Integral', `\\int{${P}}`),
			s('integralLimits', 'Integral with Limits', `\\int_{${P}}^{${P}}{${P}}`),
			s('double', 'Double Integral', `\\iint{${P}}`),
			s('doubleLimits', 'Double Integral with Limits', `\\iint_{${P}}^{${P}}{${P}}`),
			s('triple', 'Triple Integral', `\\iiint{${P}}`),
			s('tripleLimits', 'Triple Integral with Limits', `\\iiint_{${P}}^{${P}}{${P}}`),
			s('contour', 'Contour Integral', `\\oint{${P}}`),
			s('contourLimits', 'Contour Integral with Limits', `\\oint_{${P}}^{${P}}{${P}}`),
			s('dx', 'Differential x', `d${P}`),
		],
	},
	{
		id: 'largeOperator',
		name: 'Large Operator',
		structures: [
			s('sum', 'Summation', `\\sum{${P}}`),
			s('sumLimits', 'Summation with Limits', `\\sum_{${P}}^{${P}}{${P}}`),
			s('sumSub', 'Summation with Subscript Limit', `\\sum_{${P}}{${P}}`),
			s('product', 'Product', `\\prod{${P}}`),
			s('productLimits', 'Product with Limits', `\\prod_{${P}}^{${P}}{${P}}`),
			s('coproduct', 'Co-Product with Limits', `\\coprod_{${P}}^{${P}}{${P}}`),
			s('union', 'Union with Limits', `\\bigcup_{${P}}^{${P}}{${P}}`),
			s('intersection', 'Intersection with Limits', `\\bigcap_{${P}}^{${P}}{${P}}`),
			s('or', 'Logical OR with Limits', `\\bigvee_{${P}}^{${P}}{${P}}`),
			s('and', 'Logical AND with Limits', `\\bigwedge_{${P}}^{${P}}{${P}}`),
			s('binomialSum', 'Binomial Sum', '\\sum_{k=0}^{n}{\\binom{n}{k}}'),
		],
	},
	{
		id: 'bracket',
		name: 'Bracket',
		structures: [
			s('parentheses', 'Parentheses', `\\left(${P}\\right)`),
			s('brackets', 'Brackets', `\\left[${P}\\right]`),
			s('braces', 'Braces', `\\left\\{${P}\\right\\}`),
			s('angles', 'Angle Brackets', `\\left\\langle ${P}\\right\\rangle `),
			s('floor', 'Floor', `\\left\\lfloor ${P}\\right\\rfloor `),
			s('ceiling', 'Ceiling', `\\left\\lceil ${P}\\right\\rceil `),
			s('bars', 'Single Bars', `\\left|${P}\\right|`),
			s('doubleBars', 'Double Bars', `\\left\\|${P}\\right\\|`),
			s('cases', 'Cases', `\\begin{cases}${P}\\\\${P}\\end{cases}`),
			s('binomial', 'Binomial Coefficient', `\\binom{${P}}{${P}}`),
		],
	},
	{
		id: 'function',
		name: 'Function',
		structures: [
			s('sin', 'Sine Function', `\\sin{${P}}`),
			s('cos', 'Cosine Function', `\\cos{${P}}`),
			s('tan', 'Tangent Function', `\\tan{${P}}`),
			s('cot', 'Cotangent Function', `\\cot{${P}}`),
			s('sec', 'Secant Function', `\\sec{${P}}`),
			s('csc', 'Cosecant Function', `\\csc{${P}}`),
			s('asin', 'Inverse Sine Function', `\\sin^{-1}{${P}}`),
			s('acos', 'Inverse Cosine Function', `\\cos^{-1}{${P}}`),
			s('sinh', 'Hyperbolic Sine Function', `\\sinh{${P}}`),
			s('cosh', 'Hyperbolic Cosine Function', `\\cosh{${P}}`),
			s('tanh', 'Hyperbolic Tangent Function', `\\tanh{${P}}`),
			s('sinTheta', 'Sine Theta', '\\sin{\\theta}'),
			s('cos2x', 'Cosine 2x', '\\cos{2x}'),
		],
	},
	{
		id: 'accent',
		name: 'Accent',
		structures: [
			s('dot', 'Dot', `\\dot{${P}}`),
			s('ddot', 'Double Dot', `\\ddot{${P}}`),
			s('dddot', 'Triple Dot', `\\dddot{${P}}`),
			s('hat', 'Hat', `\\hat{${P}}`),
			s('check', 'Check', `\\check{${P}}`),
			s('acute', 'Acute', `\\acute{${P}}`),
			s('grave', 'Grave', `\\grave{${P}}`),
			s('breve', 'Breve', `\\breve{${P}}`),
			s('tilde', 'Tilde', `\\tilde{${P}}`),
			s('bar', 'Bar', `\\bar{${P}}`),
			s('vector', 'Rightwards Arrow Above', `\\vec{${P}}`),
			s('leftArrow', 'Leftwards Arrow Above', `\\overleftarrow{${P}}`),
			s('overbar', 'Overbar', `\\overline{${P}}`),
			s('underbar', 'Underbar', `\\underline{${P}}`),
			s('overbrace', 'Overbrace', `\\overbrace{${P}}`),
			s('underbrace', 'Underbrace', `\\underbrace{${P}}`),
			s('boxed', 'Boxed Formula', `\\boxed{${P}}`),
		],
	},
	{
		id: 'limitLog',
		name: 'Limit and Log',
		structures: [
			s('logBase', 'Logarithm', `\\log_{${P}}{${P}}`),
			s('log', 'Log', `\\log{${P}}`),
			s('limit', 'Limit', `\\lim_{${P}}{${P}}`),
			s('min', 'Minimum', `\\min_{${P}}{${P}}`),
			s('max', 'Maximum', `\\max_{${P}}{${P}}`),
			s('ln', 'Natural Logarithm', `\\ln{${P}}`),
			s('limitE', 'Limit Example', '\\lim_{n\\to\\infty}{\\left(1+\\frac{1}{n}\\right)^{n}}'),
		],
	},
	{
		id: 'operator',
		name: 'Operator',
		structures: [
			s('colonEquals', 'Colon Equal', '≔'),
			s('equalEqual', 'Equal Equal', '=='),
			s('plusEqual', 'Plus Equal', '+='),
			s('minusEqual', 'Minus Equal', '-='),
			s('deltaEqual', 'Delta Equal To', '≜'),
			s('definedAs', 'Equal to By Definition', '\\overset{\\text{def}}{=}'),
			s('measuredBy', 'Measured By', '\\overset{\\text{m}}{=}'),
			s('rightArrowAbove', 'Right Arrow with Text Above', `\\overset{${P}}{\\rightarrow}`),
			s('rightArrowBelow', 'Right Arrow with Text Below', `\\underset{${P}}{\\rightarrow}`),
			s('leftArrowAbove', 'Left Arrow with Text Above', `\\overset{${P}}{\\leftarrow}`),
			s(
				'doubleArrowAbove',
				'Left-Right Arrow with Text Above',
				`\\overset{${P}}{\\leftrightarrow}`,
			),
			s('impliesAbove', 'Implies with Text Above', `\\overset{${P}}{\\Rightarrow}`),
		],
	},
	{
		id: 'matrix',
		name: 'Matrix',
		structures: [
			s('m1x2', '1x2 Empty Matrix', `\\begin{matrix}${P}&${P}\\end{matrix}`),
			s('m2x1', '2x1 Empty Matrix', `\\begin{matrix}${P}\\\\${P}\\end{matrix}`),
			s('m1x3', '1x3 Empty Matrix', `\\begin{matrix}${P}&${P}&${P}\\end{matrix}`),
			s('m3x1', '3x1 Empty Matrix', `\\begin{matrix}${P}\\\\${P}\\\\${P}\\end{matrix}`),
			s('m2x2', '2x2 Empty Matrix', `\\begin{matrix}${P}&${P}\\\\${P}&${P}\\end{matrix}`),
			s(
				'm3x3',
				'3x3 Empty Matrix',
				`\\begin{matrix}${P}&${P}&${P}\\\\${P}&${P}&${P}\\\\${P}&${P}&${P}\\end{matrix}`,
			),
			s('identity2', '2x2 Identity Matrix', '\\begin{matrix}1&0\\\\0&1\\end{matrix}'),
			s('identity3', '3x3 Identity Matrix', '\\begin{matrix}1&0&0\\\\0&1&0\\\\0&0&1\\end{matrix}'),
			s(
				'parentheses',
				'2x2 Matrix in Parentheses',
				`\\begin{pmatrix}${P}&${P}\\\\${P}&${P}\\end{pmatrix}`,
			),
			s(
				'brackets',
				'2x2 Matrix in Brackets',
				`\\begin{bmatrix}${P}&${P}\\\\${P}&${P}\\end{bmatrix}`,
			),
			s('determinant', '2x2 Determinant', `\\begin{vmatrix}${P}&${P}\\\\${P}&${P}\\end{vmatrix}`),
		],
	},
];

/** The structure `id` of `group`, if any. */
export function equationStructure(group: string, id: string): EquationStructure | undefined {
	return EQUATION_STRUCTURES.find((gallery) => gallery.id === group)?.structures.find(
		(entry) => entry.id === id,
	);
}
