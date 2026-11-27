// The Equation tab: Word's contextual tab while the caret is in an equation. Tools (the built-in
// Equation gallery, Ink Equation, Convert between Professional and Linear, Normal Text and the
// Unicode or LaTeX input format), Symbols (the symbol gallery by category) and Structures (eleven
// galleries that insert a structure with placeholders). The galleries and the editing they drive
// are loaded on first use (`equation-galleries.ts`, `equation-editing.ts`).
import { NodeSelection, Plugin, PluginKey, type EditorState } from 'prosemirror-state';
import { group, menuSelect, stack, tool } from './ribbon-parts';
import { ribbonIcon } from './ribbon-icons';
import type { RibbonIcon } from './ribbon-icons';

export type EquationCommand =
	| 'new'
	| 'gallery'
	| 'builtIn'
	| 'structure'
	| 'symbol'
	| 'convert'
	| 'normalText'
	| 'format'
	| 'ink';

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

/** A button that opens equation gallery `value`: a menu, collapsed until it opens. */
function opensGallery(button: HTMLButtonElement, value: string): HTMLButtonElement {
	button.dataset.equationGallery = value;
	button.setAttribute('aria-haspopup', 'menu');
	button.setAttribute('aria-expanded', 'false');
	return button;
}

/**
 * Insert > Equation, a split button as in Word: the button inserts a new equation (Alt+=), its
 * arrow opens the built-in equations, Insert New Equation and the LaTeX dialog.
 */
export function equationSplit(): HTMLElement {
	const wrap = document.createElement('div');
	wrap.className = 'ribbon-split-inline ribbon-command-split ribbon-command-split-large';
	const main = tool('Equation', 'equation', { type: 'equation', key: 'new' }, { large: true });
	const caret = tool('Built-In Equations', 'caret', gallery('insert-builtIn'), {
		className: 'ribbon-menu ribbon-menu-caret dve-equation-caret',
	});
	caret.replaceChildren(ribbonIcon('caret', 12));
	wrap.append(main, opensGallery(caret, 'insert-builtIn'));
	return wrap;
}

export function buildEquationPanel(panel: HTMLElement): void {
	// Available once the host sets a handwriting recognizer (`syncEquationTab`).
	const ink = tool(
		'Ink Equation',
		'inkEquation',
		{ type: 'equation', key: 'ink' },
		{ large: true },
	);
	ink.dataset.equationInk = '';
	ink.dataset.unavailable = '';
	ink.disabled = true;
	ink.title = 'Ink Equation needs handwriting recognition, which this editor does not have.';
	const builtIn = opensGallery(
		tool('Equation', 'equation', gallery('builtIn'), { large: true }),
		'builtIn',
	);
	// The gallery and Convert apply to a read-only equation too (see `syncEquationTab`).
	builtIn.dataset.equationAny = '';
	const symbols = opensGallery(
		tool('Symbols', 'symbol', gallery('symbols'), { large: true }),
		'symbols',
	);
	const convert = menuSelect(
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
	);
	convert.dataset.equationAny = '';
	panel.classList.add('dve-equation-panel');
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
				convert,
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
			...STRUCTURE_BUTTONS.map(([value, label, icon]) =>
				opensGallery(tool(label, icon, gallery(value), { large: true }), value),
			),
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

const trackKey = new PluginKey<{ pos: number | undefined; same: boolean }>('equationTrack');

/**
 * Follows the selected equation through transactions: `same` stays true while the selection
 * remains on one equation, even when edits elsewhere (collaboration) move its position.
 */
export const equationTrackPlugin = () =>
	new Plugin({
		key: trackKey,
		state: {
			init: (_config, state) => ({ pos: selectedEquationPos(state), same: false }),
			apply(tr, value, _old, state) {
				const pos = selectedEquationPos(state);
				const mapped = value.pos === undefined ? undefined : tr.mapping.map(value.pos);
				return { pos, same: pos !== undefined && pos === mapped };
			},
		},
	});

/** Whether the selection is still on the equation it was on (see `equationTrackPlugin`). */
export function sameEquation(state: EditorState): boolean {
	return trackKey.getState(state)?.same ?? false;
}
