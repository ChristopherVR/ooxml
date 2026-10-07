import type { EditorView } from 'prosemirror-view';
import { isStHighlightColor } from 'ooxml-core/docx';
import { selectionScript } from './script-state';
import { applyRunFormattingPatch } from 'ooxml-core/docx/ui';

export function applyHighlight(view: EditorView, color: string) {
	if (!isStHighlightColor(color)) return;
	applyRunFormattingPatch({ highlight: color === 'none' ? undefined : color })(
		view.state,
		view.dispatch,
		view,
	);
}

export function toggleVerticalAlign(view: EditorView, value: 'superscript' | 'subscript') {
	applyVerticalAlign(view, selectionScript(view.state) === value ? 'baseline' : value);
}

export function applyVerticalAlign(
	view: EditorView,
	value: 'baseline' | 'superscript' | 'subscript',
) {
	applyRunFormattingPatch({ verticalAlign: value })(view.state, view.dispatch, view);
}
