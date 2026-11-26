// A press on an inline picture or equation drawn on a Print Layout line selects it (a node
// selection), as in the Draft view, so its contextual tab and commands apply to it.
import { NodeSelection } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';

/** Selects the inline picture under `target`; false when the press is not on one. */
export function selectInlineObject(
	view: EditorView,
	target: EventTarget | null,
	position: (blockId: string, offset: number) => number | undefined,
): boolean {
	const picture =
		target instanceof Element ? target.closest<HTMLElement>('.dve-print-picture') : null;
	const line = picture?.closest<HTMLElement>('.dve-print-line[data-block-id]');
	if (!picture || !line || picture.classList.contains('dve-print-float')) return false;
	const pos = position(line.dataset.blockId!, Number(picture.dataset.sourceStart));
	const node = pos === undefined ? null : view.state.doc.nodeAt(pos);
	if (pos === undefined || (node?.type.name !== 'image' && node?.type.name !== 'equation'))
		return false;
	view.dispatch(view.state.tr.setSelection(NodeSelection.create(view.state.doc, pos)));
	view.focus();
	return true;
}
