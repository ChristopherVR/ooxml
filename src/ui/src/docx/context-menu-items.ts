/** Pure context-menu model: which items appear, and which are disabled, for a given editor state. */
import type { EditorView } from 'prosemirror-view';
import { selectedChart } from 'ooxml-core/docx/ui';
import type { DrawingWrap } from 'ooxml-core/docx';
import { linkAtSelection } from './link-commands';
import { pictureWrapOf, selectedPicturePos, VML_EDIT_REASON } from './picture-format-commands';
import { noteReferenceAtSelection, type NoteKind } from './note-commands';
import type { LocalizationKey } from './localization';
import type { RibbonAction } from './ribbon-action';
import { schema } from './schema';
import {
	canExecuteTableCommand,
	tableCommandUnavailableReason,
	type TableCommand,
} from './table-commands';

export type MenuCommand =
	| { kind: 'clipboard'; op: 'cut' | 'copy' | 'paste' }
	| { kind: 'removeLink' }
	| { kind: 'ribbon'; action: RibbonAction };

export interface MenuItem {
	id: string;
	label: LocalizationKey;
	command: MenuCommand;
	disabled: boolean;
	/** Reason shown when the command is disabled. */
	reason?: LocalizationKey;
	/** Starts a new group; rendered as a separator. */
	separatorBefore: boolean;
	/** A heading shown before this row (Wrap Text over its choices). */
	heading?: LocalizationKey;
	/** A choice in this state (the picture's current wrapping). */
	checked?: boolean;
}

export interface MenuState {
	readOnly: boolean;
	hasSelection: boolean;
	onLink: boolean;
	inTable: boolean;
	/** An editable chart is selected: Edit Data and Format Chart Area appear. */
	onChart?: boolean;
	/** The Wrap Text choice of the selected picture; set only while a picture is selected. */
	pictureWrap?: DrawingWrap;
	vmlFloat?: true;
	/** The caret touches a note reference: Convert to Endnote/Footnote appears. */
	onNote?: { kind: NoteKind; id: string } | undefined;
	/** The browser exposes an async clipboard read API. */
	canReadClipboard: boolean;
	canTable(command: TableCommand): boolean;
}

const TABLE_ITEMS: [TableCommand, LocalizationKey][] = [
	['rowBefore', 'Insert row above'],
	['rowAfter', 'Insert row below'],
	['deleteRow', 'Delete row'],
	['columnBefore', 'Insert column left'],
	['columnAfter', 'Insert column right'],
	['deleteColumn', 'Delete column'],
	['mergeCells', 'Merge cells'],
	['splitCell', 'Split cells'],
	['deleteTable', 'Delete table'],
];

const WRAP_ITEMS: [DrawingWrap, LocalizationKey][] = [
	['inline', 'In Line with Text'],
	['square', 'Square'],
	['tight', 'Tight'],
	['through', 'Through'],
	['topAndBottom', 'Top and Bottom'],
	['behind', 'Behind Text'],
	['front', 'In Front of Text'],
];

/** Items for the current state; table items only exist while the selection is in a table. */
export function computeMenuItems(state: MenuState): MenuItem[] {
	const groups: Omit<MenuItem, 'separatorBefore'>[][] = [];
	const clip = (op: 'cut' | 'copy' | 'paste', label: LocalizationKey, disabled: boolean) => ({
		id: op,
		label,
		command: { kind: 'clipboard', op } as const,
		disabled,
		...(op === 'paste' && !state.canReadClipboard && !state.readOnly
			? { reason: 'menu.pasteUnavailable' as const }
			: {}),
	});
	groups.push([
		clip('cut', 'menu.cut', state.readOnly || !state.hasSelection),
		clip('copy', 'menu.copy', !state.hasSelection),
		clip('paste', 'menu.paste', state.readOnly || !state.canReadClipboard),
	]);
	const link = (id: string, label: LocalizationKey, command: MenuCommand) => ({
		id,
		label,
		command,
		disabled: state.readOnly,
	});
	const insertLink = { kind: 'ribbon', action: { type: 'link' } } as const;
	groups.push(
		state.onLink
			? [
					link('link-edit', 'menu.editLink', insertLink),
					link('link-remove', 'Remove link', { kind: 'removeLink' }),
				]
			: [link('link-insert', 'Insert link', insertLink)],
	);
	groups.push([
		link('comment-add', 'Add comment', {
			kind: 'ribbon',
			action: { type: 'comments', key: 'add' },
		}),
	]);
	if (state.onNote) {
		const { kind, id } = state.onNote;
		groups.push([
			link('note-convert', kind === 'footnote' ? 'Convert to Endnote' : 'Convert to Footnote', {
				kind: 'ribbon',
				action: { type: 'convertNotes', selection: { kind, id } },
			}),
		]);
	}
	if (state.onChart)
		groups.push([
			link('chart-data', 'Edit Data', {
				kind: 'ribbon',
				action: { type: 'formatDialog', kind: 'chartData' },
			}),
			link('chart-format', 'Format Chart Area...', {
				kind: 'ribbon',
				action: { type: 'chart', key: 'format' },
			}),
		]);
	if (state.pictureWrap)
		groups.push(
			WRAP_ITEMS.map(([wrap, label], index) => ({
				id: `picture-wrap-${wrap}`,
				label,
				command: { kind: 'ribbon', action: { type: 'picture', key: 'wrap', value: wrap } } as const,
				disabled: state.readOnly || (state.vmlFloat === true && wrap === 'inline'),
				...(state.vmlFloat && wrap === 'inline' ? { reason: VML_EDIT_REASON } : {}),
				checked: state.pictureWrap === wrap,
				...(index === 0 ? { heading: 'Wrap text' as const } : {}),
			})),
			[
				link('picture-format', 'Format picture', {
					kind: 'ribbon',
					action: { type: 'formatPicture' },
				}),
			],
		);
	if (state.inTable)
		groups.push(
			TABLE_ITEMS.map(([key, label]) => {
				const disabled = state.readOnly || !state.canTable(key);
				const reason = disabled && tableCommandUnavailableReason(key);
				return {
					id: `table-${key}`,
					label,
					command:
						key === 'splitCell'
							? ({ kind: 'ribbon', action: { type: 'formatDialog', kind: 'splitCells' } } as const)
							: ({ kind: 'ribbon', action: { type: 'tableEdit', key } } as const),
					disabled,
					...(reason ? { reason } : {}),
				};
			}),
		);
	return groups.flatMap((group, index) =>
		group.map((item, position) => ({ ...item, separatorBefore: index > 0 && position === 0 })),
	);
}

function selectionInTable(view: EditorView): boolean {
	const { $from } = view.state.selection;
	for (let depth = $from.depth; depth > 0; depth--)
		if ($from.node(depth).type === schema.nodes.table) return true;
	return false;
}

function pictureWrapState(view: EditorView): Pick<MenuState, 'pictureWrap' | 'vmlFloat'> {
	const pos = selectedPicturePos(view.state);
	const node = pos === undefined ? null : view.state.doc.nodeAt(pos);
	return node
		? { pictureWrap: pictureWrapOf(node), ...(node.attrs.vmlFloat ? { vmlFloat: true } : {}) }
		: {};
}

export function menuStateFromView(view: EditorView, readOnly: boolean): MenuState {
	return {
		readOnly: readOnly || !view.editable,
		hasSelection: !view.state.selection.empty,
		onLink: Boolean(linkAtSelection(view)),
		inTable: selectionInTable(view),
		onChart: Boolean(selectedChart(view.state)?.editable),
		...pictureWrapState(view),
		onNote: noteReferenceAtSelection(view),
		canReadClipboard:
			typeof navigator !== 'undefined' &&
			(typeof navigator.clipboard?.read === 'function' ||
				typeof navigator.clipboard?.readText === 'function'),
		canTable: (command) => canExecuteTableCommand(view, command),
	};
}
