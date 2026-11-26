/**
 * The characters of the linear format (UnicodeMath, Unicode Technical Note 28) and the math
 * AutoCorrect list that turns a control word such as `\alpha` into its character when a space
 * follows it, as Word's equation editor does.
 */
import { GREEK_MAP, NARY_MAP, OPERATOR_MAP } from './index';

/** Build-up operators and the characters that structure linear text. */
export const LINEAR = {
	fraction: '/',
	linearFraction: '∕',
	stack: '¦',
	below: '┬',
	above: '┴',
	naryand: '▒',
	matrix: '■',
	eqArray: '█',
	box: '□',
	overbar: '¯',
	underbar: '▁',
	overbrace: '⏞',
	underbrace: '⏟',
	funcApply: '⁡',
	invisibleOpen: '├',
	invisibleClose: '┤',
	groupOpen: '〖',
	groupClose: '〗',
} as const;

export const ROOTS: Record<string, string> = { '√': '', '∛': '3', '∜': '4' };

/** N-ary operator characters the converters know, keyed to their LaTeX command. */
export const NARY_CHARS: Record<string, string> = Object.fromEntries(
	Object.entries(NARY_MAP).map(([command, char]) => [char, command]),
);

/** Opening brackets and the bracket that closes each. */
export const BRACKETS: Record<string, string> = {
	'(': ')',
	'[': ']',
	'{': '}',
	'⟨': '⟩',
	'⌊': '⌋',
	'⌈': '⌉',
	'|': '|',
	'‖': '‖',
	[LINEAR.invisibleOpen]: LINEAR.invisibleClose,
};
export const CLOSERS = new Set([')', ']', '}', '⟩', '⌋', '⌉', LINEAR.invisibleClose]);

/** Function names Word sets upright and applies to the argument that follows. */
export const LINEAR_FUNCTIONS = new Set([
	'sin',
	'cos',
	'tan',
	'cot',
	'sec',
	'csc',
	'arcsin',
	'arccos',
	'arctan',
	'sinh',
	'cosh',
	'tanh',
	'coth',
	'log',
	'ln',
	'exp',
	'lim',
	'min',
	'max',
	'det',
	'gcd',
	'arg',
]);

const strip = (map: Record<string, string>) =>
	Object.fromEntries(Object.entries(map).map(([command, char]) => [command.slice(1), char]));

/** Control word (without its backslash) to character: Word's math AutoCorrect entries. */
export const MATH_AUTOCORRECT: Readonly<Record<string, string>> = {
	...strip(GREEK_MAP),
	...strip(OPERATOR_MAP),
	...strip(NARY_MAP),
	varepsilon: 'ε',
	epsilon: 'ϵ',
	sqrt: '√',
	cbrt: '∛',
	qdrt: '∜',
	matrix: LINEAR.matrix,
	eqarray: LINEAR.eqArray,
	naryand: LINEAR.naryand,
	of: LINEAR.naryand,
	box: LINEAR.box,
	overbar: LINEAR.overbar,
	underbar: LINEAR.underbar,
	overbrace: LINEAR.overbrace,
	underbrace: LINEAR.underbrace,
	below: LINEAR.below,
	above: LINEAR.above,
	atop: LINEAR.stack,
	funcapply: LINEAR.funcApply,
	degree: '°',
	cong: '≅',
	sim: '∼',
	simeq: '≃',
	ni: '∋',
	emptyset: '∅',
	hbar: 'ℏ',
	ell: 'ℓ',
	Re: 'ℜ',
	Im: 'ℑ',
	aleph: 'ℵ',
	angle: '∠',
	perp: '⊥',
	parallel: '∥',
	oplus: '⊕',
	otimes: '⊗',
	circ: '∘',
	bullet: '∙',
	ast: '∗',
	star: '⋆',
	langle: '⟨',
	rangle: '⟩',
	lfloor: '⌊',
	rfloor: '⌋',
	lceil: '⌈',
	rceil: '⌉',
	uparrow: '↑',
	downarrow: '↓',
	mapsto: '↦',
	defeq: '≝',
	coloneq: '≔',
	triangleq: '≜',
	dots: '…',
	vbar: '│',
	Vert: '‖',
	prime: '′',
	dd: 'ⅆ',
	ii: 'ⅈ',
	ee: 'ⅇ',
	hat: '̂',
	check: '̌',
	tilde: '̃',
	acute: '́',
	grave: '̀',
	dot: '̇',
	ddot: '̈',
	dddot: '⃛',
	breve: '̆',
	bar: '̅',
	vec: '⃗',
};

/** Whether `char` is a combining mark that the linear format writes after an accent's base. */
export const isCombining = (char: string) => /^[̀-ͯ⃐-⃿]$/u.test(char);

/** Whether `char` belongs to an operand run (letters, digits and placeholders). */
export const isOperandChar = (char: string) => /^[\p{L}\p{N}⬚∞]$/u.test(char);
