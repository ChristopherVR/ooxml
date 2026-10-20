// Draws each symbol character (`w:sym`) in the symbol's own font, as Word does, while the rest of
// its run keeps the run's font.
import type { Node } from 'prosemirror-model';
import { Plugin, type Transaction } from 'prosemirror-state';
import { ReplaceStep } from 'prosemirror-transform';
import { Decoration, DecorationSet } from 'prosemirror-view';
import type { TextRun } from '../model';
import { symbolCharacter } from '../special-characters';

const symbolOf = (node: Node): TextRun['symbol'] =>
	(
		node.marks.find((mark) => mark.type.name === 'runProperties')?.attrs.props as {
			symbol?: TextRun['symbol'];
		} | null
	)?.symbol;

function symbolDecorations(doc: Node): DecorationSet {
	const decorations: Decoration[] = [];
	doc.descendants((node, pos) => {
		if (!node.isText) return true;
		const symbol = symbolOf(node);
		const character = symbol && symbolCharacter(symbol.char);
		if (!symbol || !character) return false;
		const style = `font-family:"${symbol.font.replace(/["';{}\\]/g, '')}"`;
		for (
			let at = node.text!.indexOf(character);
			at >= 0;
			at = node.text!.indexOf(character, at + 1)
		)
			decorations.push(Decoration.inline(pos + at, pos + at + character.length, { style }));
		return false;
	});
	return decorations.length ? DecorationSet.create(doc, decorations) : DecorationSet.empty;
}

/** Whether the transaction brings in text carrying a symbol (a paste, a remote change). */
function insertsSymbol(tr: Transaction): boolean {
	return tr.steps.some((step) => {
		if (!(step instanceof ReplaceStep)) return false;
		let found = false;
		step.slice.content.descendants((node) => {
			if (node.isText && symbolOf(node)) found = true;
			return !found;
		});
		return found;
	});
}

export function symbolFontPlugin(): Plugin<DecorationSet> {
	return new Plugin<DecorationSet>({
		state: {
			init: (_config, state) => symbolDecorations(state.doc),
			apply: (tr, set) =>
				!tr.docChanged
					? set
					: set === DecorationSet.empty && !insertsSymbol(tr)
						? set
						: symbolDecorations(tr.doc),
		},
		props: {
			decorations(state) {
				return this.getState(state);
			},
		},
	});
}
