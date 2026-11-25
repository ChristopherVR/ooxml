import {
	canBaseStyleOn,
	stylesOf,
	type DocumentModel,
	type ParagraphFormatting,
	type RunFormatting,
	type StyleEdit,
	type StyleKind,
} from 'ooxml-core/docx';
import { twipsFromPoints } from 'ooxml-core/units';
import {
	dialogButton,
	fieldset,
	labelled,
	numberInput,
	row,
	selectOf,
	textInput,
} from './dialog-fields';
import {
	createDialogShell,
	dialogActions,
	localizeDialog,
	onDialogDismiss,
	setDialogOpen,
} from './dialog-shell';
import { translate, type EditorLocale } from './localization';
import { selectionStyle } from './style-commands';

const KINDS = [
	['paragraph', 'Paragraph'],
	['character', 'Character style'],
	['linked', 'Linked style'],
	['table', 'Table'],
] as const;
const ALIGNMENTS = [
	['', 'styles.inherit'],
	['left', 'Left'],
	['center', 'Center'],
	['right', 'Right'],
	['justify', 'Justified'],
] as const;
const TOGGLES = [
	['', 'styles.inherit'],
	['on', 'On'],
	['off', 'Off'],
] as const;
const kindOf = (value: string): StyleKind =>
	KINDS.find(([kind]) => kind === value)?.[0] ?? 'paragraph';

type StyleDialogOptions = {
	model: DocumentModel;
	selection: ReturnType<typeof selectionStyle>;
	locale: EditorLocale;
	modify: boolean;
	submit(edit: StyleEdit): boolean;
};

export function createStyleDialog(options: StyleDialogOptions) {
	const { model, selection, locale, modify, submit } = options;
	const dialog = createDialogShell(
		modify ? 'Modify Style' : 'Create a Style',
		'dve-format-dialog dve-style-dialog',
	);
	const chosen = selectOf([]);
	const name = textInput();
	name.maxLength = 253;
	const type = selectOf(KINDS);
	type.disabled = modify;
	const basedOn = selectOf([]);
	const following = selectOf([]);
	const family = textInput();
	const size = numberInput(1, 1638, 0.5);
	const bold = selectOf(TOGGLES);
	const italic = selectOf(TOGGLES);
	const align = selectOf(ALIGNMENTS);
	const before = numberInput(0, 1584, 0.05);
	const after = numberInput(0, 1584, 0.05);
	const nextRow = labelled('Style for following paragraph', following);
	const alignment = fieldset(
		'Paragraph',
		labelled('Alignment', align),
		row(labelled('Before', before), labelled('After', after)),
	);
	const error = document.createElement('p');
	error.setAttribute('role', 'alert');
	error.hidden = true;
	error.textContent = 'Enter a style name and valid formatting.';
	const ok = dialogButton('OK', true);
	const cancel = dialogButton('Cancel');
	dialog.append(
		...(modify ? [labelled('Styles', chosen)] : []),
		row(labelled('Style name', name), labelled('Style type', type)),
		labelled('Based on', basedOn),
		nextRow,
		fieldset(
			'Font',
			row(labelled('Font', family), labelled('Font size', size)),
			row(labelled('Bold', bold), labelled('Italic', italic)),
		),
		alignment,
		error,
		dialogActions(ok, cancel),
	);
	const choices = [
		...Object.values(stylesOf(model, 'paragraph')).map((style) => {
			const linked =
				model.paragraphStyles?.styles[style.id]?.linkedStyle ??
				model.characterStyles?.styles[style.id]?.linkedStyle;
			return { kind: kindOf(linked ? 'linked' : 'paragraph'), style };
		}),
		...(['character', 'table'] as const).flatMap((kind) =>
			Object.values(stylesOf(model, kind)).map((style) => ({ kind, style })),
		),
	];
	for (const { style, kind } of choices)
		chosen.add(new Option(style.name ?? style.id, kind + ':' + style.id));
	if (modify) {
		const selected =
			choices.find(
				({ kind, style }) => kind === 'character' && style.id === selection.characterId,
			) ??
			choices.find(({ kind, style }) => kind === 'table' && style.id === selection.tableId) ??
			choices.find(
				({ kind, style }) =>
					(kind === 'paragraph' || kind === 'linked') && style.id === selection.paragraphId,
			);
		if (selected) chosen.value = selected.kind + ':' + selected.style.id;
	}
	const toggleValue = (value: boolean | undefined) =>
		value === undefined ? '' : value ? 'on' : 'off';
	const toggleState = (value: string) => (value ? value === 'on' : undefined);
	let id: string | undefined;
	let run: RunFormatting = {};
	let paragraph: ParagraphFormatting = {};
	function reload() {
		const current = modify ? choices[chosen.selectedIndex] : undefined;
		const kind = current?.kind ?? kindOf(type.value);
		type.value = kind;
		id = current?.style.id;
		if (current) name.value = current.style.name ?? current.style.id;
		const styles = stylesOf(model, kind);
		const bases = Object.values(styles).filter((style) => canBaseStyleOn(styles, style.id, id));
		basedOn.replaceChildren(
			new Option(translate(locale, 'None'), ''),
			...bases.map((style) => new Option(style.name ?? style.id, style.id)),
		);
		const sourceBase =
			current?.style.basedOn ??
			(modify
				? ''
				: kind === 'character'
					? selection.characterId
					: kind === 'table'
						? selection.tableId
						: selection.paragraphId);
		if (sourceBase && modify && !bases.some((style) => style.id === sourceBase))
			basedOn.add(new Option(sourceBase, sourceBase));
		basedOn.value = sourceBase ?? '';
		if (basedOn.selectedIndex < 0) basedOn.value = '';
		const paragraphStyle =
			id && (kind === 'paragraph' || kind === 'linked')
				? model.paragraphStyles?.styles[id]
				: undefined;
		const tableStyle = id && kind === 'table' ? model.tableStyles?.styles[id] : undefined;
		const followingStyles = Object.values(stylesOf(model, 'paragraph'));
		following.replaceChildren(
			new Option(translate(locale, 'Same style'), ''),
			...followingStyles.map((style) => new Option(style.name ?? style.id, style.id)),
		);
		if (paragraphStyle?.next && !followingStyles.some((style) => style.id === paragraphStyle.next))
			following.add(new Option(paragraphStyle.next, paragraphStyle.next));
		following.value = paragraphStyle?.next ?? '';
		nextRow.hidden = kind === 'character' || kind === 'table';
		alignment.hidden = kind === 'character';
		run = modify
			? ((kind === 'table'
					? tableStyle?.runFormatting
					: id
						? model.characterStyles?.styles[id]?.formatting
						: undefined) ?? {})
			: selection.runFormatting;
		paragraph = modify
			? ((kind === 'table' ? tableStyle?.paragraphFormatting : paragraphStyle?.formatting) ?? {})
			: selection.paragraphFormatting;
		family.value = run.fontFamily ?? '';
		size.value = run.fontSize === undefined ? '' : String(run.fontSize);
		bold.value = toggleValue(run.bold);
		italic.value = toggleValue(run.italic);
		align.value = paragraph.align ?? '';
		if (align.selectedIndex < 0) align.value = '';
		before.value =
			paragraph.spacingBeforeTwips === undefined ? '' : String(paragraph.spacingBeforeTwips / 20);
		after.value =
			paragraph.spacingAfterTwips === undefined ? '' : String(paragraph.spacingAfterTwips / 20);
	}
	const close = () => {
		setDialogOpen(dialog, false);
		dialog.remove();
	};
	chosen.addEventListener('change', reload);
	type.addEventListener('change', reload);
	cancel.addEventListener('click', close);
	onDialogDismiss(dialog, close);
	ok.addEventListener('click', () => {
		if (
			!name.value.trim() ||
			!size.checkValidity() ||
			(!alignment.hidden && (!before.checkValidity() || !after.checkValidity()))
		) {
			error.hidden = false;
			return;
		}
		const runFormatting = { ...run };
		const changeFamily = family.value.trim();
		if (changeFamily !== (run.fontFamily ?? '')) {
			if (changeFamily) runFormatting.fontFamily = changeFamily;
			else delete runFormatting.fontFamily;
			if (runFormatting.fontTheme) {
				const { ascii: _ascii, hAnsi: _hAnsi, ...otherScripts } = runFormatting.fontTheme;
				if (Object.keys(otherScripts).length) runFormatting.fontTheme = otherScripts;
				else delete runFormatting.fontTheme;
			}
		}
		const changeSize = size.value ? size.valueAsNumber : undefined;
		if (changeSize === undefined) delete runFormatting.fontSize;
		else runFormatting.fontSize = changeSize;
		for (const [key, value] of [
			['bold', toggleState(bold.value)],
			['italic', toggleState(italic.value)],
		] as const) {
			if (value === undefined) delete runFormatting[key];
			else runFormatting[key] = value;
		}
		const paragraphFormatting = { ...paragraph };
		const textAlign = ALIGNMENTS.find(([value]) => align.value === value)?.[0];
		if (textAlign !== (paragraph.align ?? '')) {
			delete paragraphFormatting.justification;
			if (textAlign) paragraphFormatting.align = textAlign;
			else delete paragraphFormatting.align;
		}
		for (const [key, input] of [
			['spacingBeforeTwips', before],
			['spacingAfterTwips', after],
		] as const) {
			if (input.value) paragraphFormatting[key] = twipsFromPoints(input.valueAsNumber);
			else delete paragraphFormatting[key];
		}
		const kind = kindOf(type.value);
		const common = {
			...(id !== undefined && { id }),
			name: name.value,
			...(basedOn.value && { basedOn: basedOn.value }),
			runFormatting,
		};
		const edit: StyleEdit =
			kind === 'character'
				? { ...common, kind }
				: kind === 'table'
					? { ...common, kind, paragraphFormatting }
					: {
							...common,
							kind,
							paragraphFormatting,
							...(following.value && { next: following.value }),
						};
		if (submit(edit)) close();
		else error.hidden = false;
	});
	reload();
	localizeDialog(dialog, locale);
	return {
		element: dialog,
		open: () => {
			setDialogOpen(dialog, true);
			name.focus();
		},
	};
}
