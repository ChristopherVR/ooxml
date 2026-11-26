/**
 * The Equation tab, its galleries and the equation editor: the shared strings of `ooxml-core`
 * (`EQUATION_UI_STRINGS`), which the other products can show too, and the tab's name. They are
 * spread first, so a Word string with the same English text (a chart's Bar, SmartArt's Matrix)
 * keeps its own wording.
 */
import { EQUATION_UI_STRINGS } from 'ooxml-core/docx/ui';

export const enEquations = { ...EQUATION_UI_STRINGS.en, 'tab.equation': 'Equation' } as const;
type EquationStrings = Record<keyof typeof enEquations, string>;

export const frEquations: EquationStrings = {
	...EQUATION_UI_STRINGS.fr,
	'tab.equation': 'Équation',
};
export const deEquations: EquationStrings = { ...EQUATION_UI_STRINGS.de, 'tab.equation': 'Formel' };
export const esEquations: EquationStrings = {
	...EQUATION_UI_STRINGS.es,
	'tab.equation': 'Ecuación',
};
export const zhEquations: EquationStrings = {
	...EQUATION_UI_STRINGS['zh-CN'],
	'tab.equation': '公式',
};
