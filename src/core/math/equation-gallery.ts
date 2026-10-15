/** Built-in equations offered by Insert > Equation, as LaTeX the converter turns into Office Math. */
export interface GalleryEquation {
	/** Stable id, also the ribbon menu value. */
	id: 'area' | 'quadratic' | 'pythagorean' | 'binomial';
	/** English name; the editor translates it. */
	name: string;
	latex: string;
}

export const EQUATION_GALLERY: readonly GalleryEquation[] = [
	{ id: 'area', name: 'Area of Circle', latex: 'A = \\pi r^{2}' },
	{
		id: 'quadratic',
		name: 'Quadratic Formula',
		latex: 'x = \\frac{-b \\pm \\sqrt{b^{2} - 4ac}}{2a}',
	},
	{ id: 'pythagorean', name: 'Pythagorean Theorem', latex: 'a^{2} + b^{2} = c^{2}' },
	{
		id: 'binomial',
		name: 'Binomial Theorem',
		latex: '(x + a)^{n} = \\sum_{k=0}^{n} \\binom{n}{k} x^{k} a^{n-k}',
	},
];
