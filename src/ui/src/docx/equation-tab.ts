// The Equation tab: Word's contextual tab while the caret is in an equation. Tools (the built-in
// Equation gallery, Ink Equation, Convert between Professional and Linear, Normal Text and the
// Unicode or LaTeX input format), Symbols (the symbol gallery by category) and Structures (eleven
// galleries that insert a structure with placeholders). The galleries and the editing they drive
// are loaded on first use (`equation-galleries.ts`, `equation-editing.ts`).
import { NodeSelection, type EditorState } from 'prosemirror-state';
import { group, menuSelect, stack, tool } from './ribbon-parts';
import type { RibbonIcon } from './ribbon-icons';

export type EquationCommand =
	| 'new'
	| 'gallery'
	| 'builtIn'
	| 'structure'
	| 'symbol'
	| 'convert'
	| 'normalText'
	| 'format';

/** Structure galleries in Word's order, with their English labels and icons. */
export const STRUCTURE_BUTTONS: ReadonlyArray<[string, string, RibbonIcon]> = [
	['fraction', 'Fraction', 'fraction'],
	['script', 'Script', 'script'],
	['radical', 'Radical', 'radical'],
	['integral', 'Integral', 'integral'],
	['largeOperator', 'Large Operator', 'equation'],
	['bracket', 'Bracket', 'bracket'],
	['function', 'Function', 'mathFunction'],
	['accent', 'Accent', 'accent'],
	['limitLog', 'Limit and Log', 'limitLog'],
	['operator', 'Operator', 'mathOperator'],
	['matrix', 'Matrix', 'matrix'],
];

const gallery = (value: string) => ({ type: 'equation', key: 'gallery', value }) as const;

export function buildEquationPanel(panel: HTMLElement): void {
	const ink = tool('Ink Equation', 'inkEquation', null, { large: true });
	ink.title = 'Ink Equation needs handwriting recognition, which this editor does not have.';
	const builtIn = tool('Equation', 'equation', gallery('builtIn'), { large: true });
	builtIn.dataset.equationGallery = 'builtIn';
	builtIn.setAttribute('aria-haspopup', 'menu');
	const symbols = tool('Symbols', 'symbol', gallery('symbols'), { large: true });
	symbols.dataset.equationGallery = 'symbols';
	symbols.setAttribute('aria-haspopup', 'menu');
	const format = (value: 'unicode' | 'latex', label: string) => {
		const button = tool(
			label,
			value === 'unicode' ? 'symbol' : 'equation',
			{ type: 'equation', key: 'format', value },
			{ inline: true, className: 'dve-equation-format' },
		);
		button.dataset.equationFormat = value;
		button.setAttribute('aria-pressed', String(value === 'unicode'));
		return button;
	};
	panel.append(
		group(
			'Tools',
			builtIn,
			ink,
			stack(
				menuSelect(
					'Convert',
					'changeCase',
					[
						['currentProfessional', 'Current - Professional'],
						['currentLinear', 'Current - Linear'],
						['allProfessional', 'All - Professional'],
						['allLinear', 'All - Linear'],
					],
					(value) => ({ type: 'equation', key: 'convert', value }),
					{ momentary: true, inline: true },
				),
				tool(
					'Normal Text',
					'normalText',
					{ type: 'equation', key: 'normalText' },
					{ inline: true, className: 'dve-equation-normal-text' },
				),
			),
			stack(format('unicode', 'Unicode'), format('latex', 'LaTeX')),
		),
		group('Symbols', symbols),
		group(
			'Structures',
			...STRUCTURE_BUTTONS.map(([value, label, icon]) => {
				const button = tool(label, icon, gallery(value), { large: true });
				button.dataset.equationGallery = value;
				button.setAttribute('aria-haspopup', 'menu');
				return button;
			}),
		),
	);
}

/** The position of the equation the selection is on, if it is on one. */
export function selectedEquationPos(state: EditorState): number | undefined {
	const { selection } = state;
	return selection instanceof NodeSelection && selection.node.type.name === 'equation'
		? selection.from
		: undefined;
}
