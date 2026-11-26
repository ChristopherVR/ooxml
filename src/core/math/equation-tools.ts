/**
 * What the products' equation editors need once an equation is being edited: the Equation tab's
 * structure and symbol galleries with their names, the linear format (UnicodeMath and LaTeX build
 * up, Linear display) and math AutoCorrect. Loaded on first use through `loadEquationTools`.
 */
export {
	EQUATION_STRUCTURES,
	equationStructure,
	type EquationStructure,
	type EquationStructureGallery,
	type EquationStructureGroup,
} from './equation-structures';
export {
	EQUATION_SYMBOLS,
	type EquationSymbolCategory,
	type EquationSymbolCategoryId,
} from './equation-symbols';
export { EQUATION_GALLERY_STRINGS, type EquationGalleryStrings } from './equation-gallery-strings';
export {
	applyControlWords,
	buildUpEquation,
	emptyEquationXml,
	linearDisplayText,
	linearEquationXml,
	linearizeEquation,
	mathAutoCorrect,
	type EquationInputFormat,
} from './linear-format';
export { unicodeMathToLatex } from './unicode-math';
export { convertOmmlToUnicodeMath } from './omml-to-unicode-math';
export { MATH_AUTOCORRECT } from './unicode-math-symbols';
export { EQUATION_GALLERY } from './equation-gallery';
export { EQUATION_PLACEHOLDER } from './equation-placeholder';
