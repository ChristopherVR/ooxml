import { resolveRunFormatting } from '../run-formatting';
import type { TextRun, DocumentModel } from '../model';
import { toggleMark } from 'prosemirror-commands';
import type { Node as ProseMirrorNode, Schema } from 'prosemirror-model';
import type { Command, EditorState, Transaction } from 'prosemirror-state';
import { appendInlineNode, inlineNodeRun, runToInlineNodes } from './run-adapter';
import { RUN_FORMAT_MARKS, setInlineRunFormatting } from './inline-formatting';

export type ToggleKey = 'bold' | 'italic' | 'underline' | 'strike';

interface Piece {
	from: number;
	to: number;
	node: ProseMirrorNode;
	paragraphStyle?: string;
}

function selectedInline(state: EditorState): Piece[] {
	const pieces: Piece[] = [];
	const { from, to } = state.selection;
	state.doc.nodesBetween(from, to, (node, pos, parent) => {
		if (!node.isInline || node.type.name === 'equation') return true;
		pieces.push({
			from: Math.max(pos, from),
			to: Math.min(pos + node.nodeSize, to),
			node,
			...(parent?.attrs.style ? { paragraphStyle: String(parent.attrs.style) } : {}),
		});
		return false;
	});
	return pieces;
}

/** Whether a piece of text shows `key`, from its marks, explicit offs and inherited styles. */
function isOn(
	state: EditorState,
	piece: Piece,
	key: ToggleKey,
	getModel: (state: EditorState) => DocumentModel | undefined,
): boolean {
	const runs: TextRun[] = [];
	appendInlineNode(runs, piece.node);
	const run = runs[0];
	if (!run) return false;
	if (run[key] !== undefined) return Boolean(run[key]);
	const model = getModel(state);
	if (!model) return false;
	return Boolean(
		resolveRunFormatting(run, {
			runCatalog: model.characterStyles,
			paragraphCatalog: model.paragraphStyles,
			paragraphStyleId: piece.paragraphStyle,
		})[key],
	);
}

/** Sets or clears the explicit off for `key` in a piece's opaque run-properties mark. */
function setExplicitOff(
	schema: Schema,
	tr: Transaction,
	piece: Piece,
	key: ToggleKey,
	off: boolean,
): void {
	const type = schema.marks.runProperties!;
	const current = type.isInSet(piece.node.marks);
	const props = { ...((current?.attrs.props as Record<string, unknown> | null) ?? {}) };
	if (off) props[key] = false;
	else delete props[key];
	tr.removeMark(piece.from, piece.to, type);
	if (Object.keys(props).length) tr.addMark(piece.from, piece.to, type.create({ props }));
}

/**
 * Word's Bold/Italic/Underline/Strikethrough: toggles what the text shows, including formatting
 * inherited from its styles. Turning off text that is on only through a style (a bold heading)
 * writes an explicit off, as Word does; turning it on clears that off again.
 */
export function createToggleFormat(
	schema: Schema,
	getModel: (state: EditorState) => DocumentModel | undefined = () => undefined,
): (key: ToggleKey) => Command {
	return (key) => {
		const markType = schema.marks[key]!;
		const plain = toggleMark(markType);
		return (state, dispatch) => {
			if (state.selection.empty) return plain(state, dispatch);
			const pieces = selectedInline(state);
			if (!pieces.length) return plain(state, dispatch);
			const allOn = pieces.every((piece) => isOn(state, piece, key, getModel));
			if (!dispatch) return true;
			const tr = state.tr;
			for (const piece of pieces) {
				if (!piece.node.isText && piece.node.type.name !== 'hardBreak') {
					const run = inlineNodeRun(piece.node)!;
					delete run[key];
					const without = runToInlineNodes(run, schema)[0]!;
					const inherited = isOn(
						state,
						{
							...piece,
							node: without.mark(
								piece.node.marks.filter((mark) => !RUN_FORMAT_MARKS.has(mark.type.name)),
							),
						},
						key,
						getModel,
					);
					if (!allOn) run[key] = true;
					else if (inherited) run[key] = false;
					setInlineRunFormatting(tr, piece.from, run);
					continue;
				}
				if (allOn) {
					tr.removeMark(piece.from, piece.to, markType);
					// Still on through a style once the mark is gone: cancel it explicitly.
					const withoutMark = {
						...piece,
						node: piece.node.mark(markType.removeFromSet(piece.node.marks)),
					};
					setExplicitOff(schema, tr, withoutMark, key, false);
					const inherited = isOn(
						state,
						{
							...withoutMark,
							node: withoutMark.node.mark(
								schema.marks.runProperties!.removeFromSet(withoutMark.node.marks),
							),
						},
						key,
						getModel,
					);
					if (inherited) setExplicitOff(schema, tr, withoutMark, key, true);
				} else {
					setExplicitOff(schema, tr, piece, key, false);
					tr.addMark(piece.from, piece.to, markType.create());
				}
			}
			dispatch(tr.scrollIntoView());
			return true;
		};
	};
}
