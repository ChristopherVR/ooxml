// Go To (names and references) and Go To Special (cells of a kind).
import { type CellRange, parseRange, sheetByName } from '@christophervr/xlsx-core';
import { type SpecialKind, selectSpecial } from '../commands/select-special.js';
import type { EditorContext } from '../context.js';
import { checkbox, field, invalid, listBox, radios, text, textInput } from './fields.js';
import { button, showDialog } from './frame.js';

/** Selects a range on a sheet, switching sheets when needed. */
export function selectOn(ctx: EditorContext, sheet: number, range: CellRange): void {
	if (sheet !== ctx.activeSheet()) ctx.setActiveSheet(sheet);
	ctx.selection.set({ sheet, ranges: [range], anchor: range.start, active: range.start });
	ctx.grid()?.scrollTo(range.start);
}

/** Resolves `A1`, `A1:B2`, `Sheet!A1`, `'My Sheet'!$A$1` or a defined name to a sheet and range. */
export function resolveReference(
	ctx: EditorContext,
	input: string,
	depth = 0,
): { sheet: number; range: CellRange } | undefined {
	const workbook = ctx.workbook();
	if (!workbook || depth > 4) return undefined;
	const ref = input.trim().replace(/^=/, '');
	if (!ref) return undefined;
	const bang = ref.lastIndexOf('!');
	let sheet = ctx.activeSheet();
	let local = ref;
	if (bang > 0) {
		const name = ref
			.slice(0, bang)
			.replace(/^'(.*)'$/, '$1')
			.replace(/''/g, "'");
		const ws = sheetByName(workbook, name);
		if (!ws) return undefined;
		sheet = workbook.sheets.indexOf(ws);
		local = ref.slice(bang + 1);
	}
	const range = parseRange(local.replace(/\$/g, ''));
	if (range) return { sheet, range };
	if (bang > 0) return undefined;
	const lower = ref.toLowerCase();
	const named =
		workbook.definedNames.find(
			(n) => n.name.toLowerCase() === lower && n.localSheet === ctx.activeSheet(),
		) ??
		workbook.definedNames.find((n) => n.name.toLowerCase() === lower && n.localSheet === undefined);
	return named ? resolveReference(ctx, named.formula, depth + 1) : undefined;
}

export function openGoTo(ctx: EditorContext): Promise<string | undefined> {
	const workbook = ctx.workbook();
	const reference = textInput(ctx);
	let confirm: (() => void) | undefined;
	const list = listBox(
		ctx,
		'Go to:',
		(name) => (reference.value = name),
		() => confirm?.(),
	);
	const names = (workbook?.definedNames ?? []).filter(
		(n) => !n.hidden && !n.name.startsWith('_xlnm.'),
	);
	list.setItems(names.map((n) => [n.name, n.name] as const));
	const special = button(ctx, 'Special...');
	return showDialog<string>(
		ctx,
		{
			name: 'go-to',
			heading: 'Go To',
			body: [
				text(ctx, 'Go to:', 'xve-field-label'),
				list.element,
				field(ctx, 'Reference:', reference),
			],
			buttons: [special],
			opened: () => reference.focus(),
			submit: () => {
				const target = resolveReference(ctx, reference.value);
				if (!target) return invalid(ctx, reference, 'Reference is not valid.');
				selectOn(ctx, target.sheet, target.range);
				return reference.value.trim();
			},
		},
		(open) => {
			confirm = () => open.element.querySelector<HTMLButtonElement>('.xve-btn-primary')?.click();
			special.addEventListener('click', () => {
				open.close();
				void ctx.dialogs.open('go-to-special');
			});
		},
	);
}

const SPECIAL: ReadonlyArray<readonly [SpecialKind, string]> = [
	['comments', 'Notes'],
	['constants', 'Constants'],
	['formulas', 'Formulas'],
	['blanks', 'Blanks'],
	['currentRegion', 'Current region'],
	['lastCell', 'Last cell'],
	['conditional', 'Conditional formats'],
	['validation', 'Data validation'],
];

const TYPES = [
	['numbers', 'Numbers'],
	['text', 'Text'],
	['logicals', 'Logicals'],
	['errors', 'Errors'],
] as const;

export function openGoToSpecial(ctx: EditorContext): Promise<SpecialKind | undefined> {
	const choice = radios(ctx, 'Select', SPECIAL, 'comments');
	const types = TYPES.map(([value, label]) => ({ value, ...checkbox(ctx, label, true) }));
	const typeBox = ctx.host.ownerDocument.createElement('div');
	typeBox.className = 'xve-row';
	typeBox.append(...types.map((t) => t.wrapper));
	const sync = (): void => {
		const kind = choice.get();
		for (const t of types) t.input.disabled = kind !== 'formulas' && kind !== 'constants';
	};
	for (const input of choice.inputs) input.addEventListener('change', sync);
	sync();
	return showDialog<SpecialKind>(ctx, {
		name: 'go-to-special',
		heading: 'Go To Special',
		body: [choice.element, typeBox],
		opened: () => choice.inputs[0]?.focus(),
		submit: () => {
			const kind = choice.get() as SpecialKind;
			const chosen = types.filter((t) => t.input.checked).map((t) => t.value);
			const query =
				kind === 'formulas' || kind === 'constants' ? { kind, types: chosen } : { kind };
			if (!selectSpecial(ctx, query)) {
				ctx.toast(ctx.t('No cells were found.'), 'info');
				return undefined;
			}
			return kind;
		},
	});
}
