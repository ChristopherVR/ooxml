/** Built-in equations offered by Insert > Equation, as LaTeX the converter turns into Office Math. */
export interface GalleryEquation {
	/** Stable id, also the ribbon menu value. */
	id:
		| 'area'
		| 'binomial'
		| 'expansion'
		| 'fourier'
		| 'pythagorean'
		| 'quadratic'
		| 'taylor'
		| 'trig1'
		| 'trig2';
	/** English name; the editor translates it. */
	name: string;
	latex: string;
}

/** Word's built-in equations, in its gallery order. */
export const EQUATION_GALLERY: readonly GalleryEquation[] = [
	{ id: 'area', name: 'Area of Circle', latex: 'A = \\pi r^{2}' },
	{
		id: 'binomial',
		name: 'Binomial Theorem',
		latex: '(x + a)^{n} = \\sum_{k=0}^{n} \\binom{n}{k} x^{k} a^{n-k}',
	},
	{
		id: 'expansion',
		name: 'Expansion of a Sum',
		latex: '(1 + x)^{n} = 1 + \\frac{nx}{1!} + \\frac{n(n-1)x^{2}}{2!} + \\cdots',
	},
	{
		id: 'fourier',
		name: 'Fourier Series',
		latex:
			'f(x) = a_{0} + \\sum_{n=1}^{\\infty}{\\left(a_{n}\\cos{\\frac{n\\pi x}{L}} + b_{n}\\sin{\\frac{n\\pi x}{L}}\\right)}',
	},
	{ id: 'pythagorean', name: 'Pythagorean Theorem', latex: 'a^{2} + b^{2} = c^{2}' },
	{
		id: 'quadratic',
		name: 'Quadratic Formula',
		latex: 'x = \\frac{-b \\pm \\sqrt{b^{2} - 4ac}}{2a}',
	},
	{
		id: 'taylor',
		name: 'Taylor Expansion',
		latex:
			'e^{x} = 1 + \\frac{x}{1!} + \\frac{x^{2}}{2!} + \\frac{x^{3}}{3!} + \\cdots, \\quad -\\infty < x < \\infty',
	},
	{
		id: 'trig1',
		name: 'Trig Identity 1',
		latex:
			'\\sin{\\alpha} \\pm \\sin{\\beta} = 2\\sin{\\frac{1}{2}}\\left(\\alpha \\pm \\beta\\right)\\cos{\\frac{1}{2}}\\left(\\alpha \\mp \\beta\\right)',
	},
	{
		id: 'trig2',
		name: 'Trig Identity 2',
		latex:
			'\\cos{\\alpha} + \\cos{\\beta} = 2\\cos{\\frac{1}{2}}\\left(\\alpha + \\beta\\right)\\cos{\\frac{1}{2}}\\left(\\alpha - \\beta\\right)',
	},
];
