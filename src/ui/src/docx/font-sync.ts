import { resolveRunFormatting } from 'ooxml-core/docx';
import type { Node as ProseMirrorNode } from 'prosemirror-model';
import type { EditorState } from 'prosemirror-state';
import { findLocalizedControl } from './localization';
import { setComboValue, type ComboInput } from './ribbon-combo';
import { runOf, styleModelOf, themeFontOf } from './run-styles';
import { schema } from './schema';
import { tableStyleAt } from './styled-paragraphs';

const DEFAULT_FAMILY = 'Calibri';
const DEFAULT_SIZE = 11;

/** The effective font family and size at a text position in the editing view. */
export function effectiveFont(
	state: EditorState,
	text: ProseMirrorNode,
	paragraph: ProseMirrorNode,
	pos: number,
): { family: string; size: number } {
	const model = styleModelOf(state);
	const run = runOf(text);
	if (!run || !model) {
		const font = text.marks.find((mark) => mark.type === schema.marks.font);
		return {
			family: String(run?.fontFamily ?? font?.attrs.family ?? DEFAULT_FAMILY),
			size: Number(run?.fontSize ?? font?.attrs.size ?? DEFAULT_SIZE),
		};
	}
	const tableStyleRun = tableStyleAt(state.doc, pos, model)?.run;
	const resolved = resolveRunFormatting(run, {
		runCatalog: model.characterStyles,
		paragraphCatalog: model.paragraphStyles,
		paragraphStyleId: paragraph.attrs.style || undefined,
		...(tableStyleRun && { tableStyleRun }),
		overrideTableStyleFontSizeAndJustification: model.overrideTableStyleFontSizeAndJustification,
	});
	return {
		family: themeFontOf(resolved, model.theme) ?? resolved.fontFamily ?? DEFAULT_FAMILY,
		size: resolved.fontSize ?? DEFAULT_SIZE,
	};
}

/** Shows the selection's font family and size in the two combo boxes (blank when they differ). */
export function syncFontControls(toolbar: HTMLElement, state: EditorState): void {
	const families = new Set<string>();
	const sizes = new Set<number>();
	const add = (text: ProseMirrorNode, paragraph: ProseMirrorNode, pos: number) => {
		const { family, size } = effectiveFont(state, text, paragraph, pos);
		families.add(family);
		sizes.add(size);
	};
	const { selection, doc } = state;
	if (selection.empty) {
		const marks = state.storedMarks ?? selection.$from.marks();
		add(schema.text('x', marks), selection.$from.parent, selection.$from.pos);
	} else {
		let found = false;
		for (const { $from, $to } of selection.ranges)
			doc.nodesBetween($from.pos, $to.pos, (node, pos, parent) => {
				if (!node.isInline || node.type.name === 'equation' || !parent) return;
				found = true;
				add(node, parent, pos);
			});
		if (!found)
			add(schema.text('x', selection.$from.marks()), selection.$from.parent, selection.$from.pos);
	}
	const family = findLocalizedControl<ComboInput>(toolbar, 'Font family');
	const size = findLocalizedControl<ComboInput>(toolbar, 'Font size');
	if (family?.comboItems)
		setComboValue(family, families.size === 1 ? ([...families][0] ?? null) : null);
	if (size?.comboItems) setComboValue(size, sizes.size === 1 ? String([...sizes][0]) : null);
}
