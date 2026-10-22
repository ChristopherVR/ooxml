import type { EditorCore } from './editor-core';
import type { RibbonAction } from './ribbon';
import { routeRibbonAction } from './ribbon-router';
import { effectiveFont } from './font-sync';
import { schema } from './schema';
import { listKindOf } from './list-commands';
import { translate } from './localization';
import { undo, redo } from 'prosemirror-history';
import { wordYjsPluginKey, wordYjsUndo, wordYjsRedo } from 'ooxml-core/docx/ui';
import { wordHostActionProps } from './building-blocks-actions';

export type WordAlignment = 'left' | 'center' | 'right' | 'justify';
/** The list kinds a control can apply; `remove` takes a paragraph out of its list. */
export type WordListChoice = 'bullet' | 'number';
/** The list the selection is in: a kind a control can apply, `other` for custom lists, or none. */
export type WordListState = WordListChoice | 'other' | null;

const ALIGNMENTS: Record<string, WordAlignment> = {
	left: 'left',
	start: 'left',
	center: 'center',
	right: 'right',
	end: 'right',
	justify: 'justify',
	both: 'justify',
};

const TOOLBAR_LABELS = [
	'Bold',
	'Italic',
	'Underline',
	'Undo',
	'Redo',
	'Font',
	'Font size',
	'Align left',
	'Align center',
	'Align right',
	'Justify',
	'Bulleted list',
	'Numbered list',
	'Insert link',
	'Find',
	'Insert picture',
] as const;

function listState(kind: ReturnType<typeof listKindOf>): WordListState {
	return kind === 'decimal' ? 'number' : kind;
}

/** A snapshot of the current editor selection, with commands bound to the same editor. */
export function wordToolbarProps(core: EditorCore) {
	const view = core.targetView();
	const state = view?.state;
	const canEdit = Boolean(view?.editable && core.canEditBody() && !core.loading);
	const values = new Map<string, Set<string | number | boolean>>();
	const add = (key: string, value: string | number | boolean) => {
		if (!values.has(key)) values.set(key, new Set());
		values.get(key)!.add(value);
	};
	if (state) {
		const sample = (
			text: import('prosemirror-model').Node,
			paragraph: import('prosemirror-model').Node,
		) => {
			for (const key of ['bold', 'italic', 'underline'])
				add(
					key,
					text.marks.some((mark) => mark.type.name === key),
				);
			const font = effectiveFont(state, text, paragraph);
			add('fontFamily', font.family);
			add('fontSize', font.size);
			add('alignment', ALIGNMENTS[String(paragraph.attrs.align ?? 'left')] ?? 'left');
		};
		if (state.selection.empty)
			sample(
				schema.text('x', state.storedMarks ?? state.selection.$from.marks()),
				state.selection.$from.parent,
			);
		else
			state.doc.nodesBetween(state.selection.from, state.selection.to, (node, _pos, parent) => {
				if (node.isText && parent) sample(node, parent);
			});
	}
	const value = <T extends string | number | boolean>(key: string, fallback: T): T | null => {
		const set = values.get(key);
		return set?.size ? (set.size === 1 ? ([...set][0] as T) : null) : fallback;
	};
	const run = (action: RibbonAction): boolean => {
		if (core.loading || !core.canEditBody() || !core.targetView()?.editable) return false;
		routeRibbonAction(core, action);
		core.refreshControls();
		return true;
	};
	const shared = state && wordYjsPluginKey.getState(state);
	return {
		labels: Object.fromEntries(
			TOOLBAR_LABELS.map((label) => [label, translate(core.locale, label)]),
		),
		canEdit,
		list: listState(
			state ? listKindOf(state.selection.$from.parent, core.model.numberingCatalog) : null,
		),
		selection: state ? { from: state.selection.from, to: state.selection.to } : null,
		bold: value<boolean>('bold', false),
		italic: value<boolean>('italic', false),
		underline: value<boolean>('underline', false),
		fontFamily: value<string>('fontFamily', 'Calibri'),
		fontSize: value<number>('fontSize', 11),
		alignment: value<WordAlignment>('alignment', 'left'),
		canUndo: Boolean(canEdit && state && (shared ? wordYjsUndo : undo)(state)),
		canRedo: Boolean(canEdit && state && (shared ? wordYjsRedo : redo)(state)),
		onBold: () => run({ type: 'format', key: 'bold' }),
		onItalic: () => run({ type: 'format', key: 'italic' }),
		onUnderline: () => run({ type: 'format', key: 'underline' }),
		onFontFamilyChange: (family: string) => run({ type: 'font', key: 'family', value: family }),
		onFontSizeChange: (size: number) => run({ type: 'font', key: 'size', value: String(size) }),
		onAlignmentChange: (alignment: WordAlignment) => run({ type: 'align', value: alignment }),
		onListChange: (kind: WordListChoice | 'remove') => run({ type: 'list', key: kind }),
		onUndo: () => run({ type: 'history', key: 'undo' }),
		onRedo: () => run({ type: 'history', key: 'redo' }),
		...wordHostActionProps(core),
	};
}
export type WordToolbarProps = ReturnType<typeof wordToolbarProps>;
