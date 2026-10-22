// Draws each symbol character (`w:sym`) in the symbol's own font, as Word does, while the rest of
// its run keeps the run's font. The decorations follow each change, and only the text nodes a
// change touched are scanned again.
import type { Node } from 'prosemirror-model';
import { Plugin, type Transaction } from 'prosemirror-state';
import { AddMarkStep, RemoveMarkStep } from 'prosemirror-transform';
import { Decoration, DecorationSet } from 'prosemirror-view';
import type { TextRun } from '../model';
import { symbolCharacter } from '../special-characters';

const symbolOf = (node: Node): TextRun['symbol'] =>
	(
		node.marks.find((mark) => mark.type.name === 'runProperties')?.attrs.props as {
			symbol?: TextRun['symbol'];
		} | null
	)?.symbol;

/** The decorations for the symbol characters of the text node at `pos`. */
function nodeDecorations(node: Node, pos: number): Decoration[] {
	const symbol = symbolOf(node);
	const character = symbol && symbolCharacter(symbol.char);
	if (!symbol || !character) return [];
	const style = `font-family:"${symbol.font.replace(/["';{}\\]/g, '')}"`;
	const decorations: Decoration[] = [];
	for (let at = node.text!.indexOf(character); at >= 0; at = node.text!.indexOf(character, at + 1))
		decorations.push(Decoration.inline(pos + at, pos + at + character.length, { style }));
	return decorations;
}

/** Recomputes the decorations of every text node between `from` and `to`. */
function rescan(set: DecorationSet, doc: Node, from: number, to: number): DecorationSet {
	let next = set;
	doc.nodesBetween(from, to, (node, pos) => {
		if (!node.isText) return true;
		const end = pos + node.nodeSize;
		const stale = next.find(pos, end).filter((item) => item.from >= pos && item.to <= end);
		next = next.remove(stale).add(doc, nodeDecorations(node, pos));
		return false;
	});
	return next;
}

/** The ranges of the new document a transaction changed: replaced content and restyled text. */
function changedRanges(tr: Transaction): [number, number][] {
	const ranges: [number, number][] = [];
	tr.steps.forEach((step, index) => {
		const later = tr.mapping.slice(index + 1);
		const add = (from: number, to: number) => ranges.push([later.map(from, -1), later.map(to, 1)]);
		if (step instanceof AddMarkStep || step instanceof RemoveMarkStep) add(step.from, step.to);
		else step.getMap().forEach((_oldStart, _oldEnd, start, end) => add(start, end));
	});
	return ranges;
}

export function symbolFontPlugin(): Plugin<DecorationSet> {
	return new Plugin<DecorationSet>({
		state: {
			init: (_config, state) => rescan(DecorationSet.empty, state.doc, 0, state.doc.content.size),
			apply(tr, set) {
				if (!tr.docChanged) return set;
				let next = set.map(tr.mapping, tr.doc);
				const size = tr.doc.content.size;
				// One position either side, so text joined to a neighbour by a deletion is seen too.
				for (const [from, to] of changedRanges(tr))
					next = rescan(next, tr.doc, Math.max(0, from - 1), Math.min(size, to + 1));
				return next;
			},
		},
		props: {
			decorations(state) {
				return this.getState(state);
			},
		},
	});
}
