import { closeHistory } from 'prosemirror-history';
import { NodeSelection, type EditorState } from 'prosemirror-state';
import type { EditorView } from 'prosemirror-view';
import { simpleTextBoxDrawing } from 'ooxml-core/docx';
import { schema } from './schema';
import type { WordArtStyle } from 'ooxml-core/drawingml';
import { normalizeWordArtGradient, wordArtRgb } from 'ooxml-core/drawingml';

export interface TextBoxSettings {
	lines: string[];
	widthPx: number;
	heightPx: number;
	border: boolean;
	wordArt?: WordArtStyle;
}

const lines = (value: unknown): string[] => {
	if (typeof value !== 'string') return [];
	try {
		const parsed: unknown = JSON.parse(value);
		return Array.isArray(parsed) ? parsed.map(String) : [];
	} catch {
		return [];
	}
};

/** The selected text box when it is the simple, editable form (inserted here or written the same way). */
export function selectedTextBox(
	state: EditorState,
): { pos: number; settings: TextBoxSettings } | null {
	const { selection } = state;
	if (!(selection instanceof NodeSelection) || selection.node.type !== schema.nodes.image)
		return null;
	const attrs = selection.node.attrs;
	if (!attrs.textBoxEditable) return null;
	const style = parsed(attrs.textBoxContent)?.wordArt;
	const fill = wordArtRgb(style?.fill);
	const outline = wordArtRgb(style?.outline);
	const shadow = wordArtRgb(style?.shadow);
	const gradient = normalizeWordArtGradient(style?.gradient);
	return {
		pos: selection.from,
		settings: {
			lines: lines(attrs.textBoxText),
			widthPx: Number(attrs.widthPx),
			heightPx: Number(attrs.heightPx),
			border: attrs.textBoxBorder !== false,
			...(fill
				? {
						wordArt: {
							fill,
							...(gradient ? { gradient } : {}),
							...(outline ? { outline } : {}),
							...(shadow ? { shadow } : {}),
							...(style?.warp ? { warp: style.warp } : {}),
						},
					}
				: {}),
		},
	};
}

const parsed = (value: unknown) => {
	if (typeof value !== 'string') return undefined;
	try {
		return JSON.parse(value);
	} catch {
		return undefined;
	}
};

/**
 * A simple text box's frame and paragraphs after its lines or outline change, as node attributes:
 * drawn like any shape, and what saving writes (`simpleTextBoxDrawing`).
 */
export function textBoxDrawingAttrs(
	nextLines: readonly string[],
	border: boolean,
	attrs: Record<string, unknown> = {},
	wordArt?: WordArtStyle,
): { shape: string; textBoxContent: string } {
	const shape = parsed(attrs.shape);
	const content = parsed(attrs.textBoxContent);
	const drawing = simpleTextBoxDrawing(nextLines, border, {
		...(shape ? { shape } : {}),
		...(content || wordArt
			? { textBoxContent: { blocks: [], ...content, ...(wordArt ? { wordArt } : {}) } }
			: {}),
		...(wordArt ? {} : { textBoxText: lines(attrs.textBoxText) }),
		textBoxBorder: attrs.textBoxBorder !== false,
	});
	return {
		shape: JSON.stringify(drawing.shape),
		textBoxContent: JSON.stringify(drawing.textBoxContent),
	};
}

const attrsFor = (settings: TextBoxSettings, attrs?: Record<string, unknown>) => ({
	unsupported: 'Text box',
	textBoxText: JSON.stringify(settings.lines),
	...textBoxDrawingAttrs(settings.lines, settings.border, attrs, settings.wordArt),
	textBoxEditable: true,
	textBoxBorder: settings.border,
	widthPx: settings.widthPx,
	heightPx: settings.heightPx,
});

/** Insert > Text Box: an inline text box at the selection, as one undo step. */
export function insertTextBox(view: EditorView, settings: TextBoxSettings): boolean {
	if (!view.editable) return false;
	const node = schema.nodes.image!.create(attrsFor(settings));
	view.dispatch(closeHistory(view.state.tr.replaceSelectionWith(node, false)).scrollIntoView());
	return true;
}

/** Applies new text, size and outline to the selected text box, as one undo step. */
export function updateTextBox(view: EditorView, pos: number, settings: TextBoxSettings): boolean {
	const node = view.state.doc.nodeAt(pos);
	if (!view.editable || !node || node.type !== schema.nodes.image || !node.attrs.textBoxEditable)
		return false;
	const tr = view.state.tr.setNodeMarkup(pos, undefined, {
		...node.attrs,
		...attrsFor(settings, node.attrs),
	});
	view.dispatch(closeHistory(tr.setSelection(NodeSelection.create(tr.doc, pos))));
	return true;
}
