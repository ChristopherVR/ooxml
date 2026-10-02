// New Name / Edit Name: name, scope, comment and refers-to, validated with the core rules.
import {
	type DefinedName,
	formatAddress,
	quoteSheetName,
	validateDefinedName,
} from '@christophervr/xlsx-core';
import type { EditorContext } from '../context.js';
import { field, invalid, select, textArea, textInput } from './fields.js';
import { showDialog } from './frame.js';

export interface DefineNameProps {
	name?: string;
	localSheet?: number;
}

const absolute = (row: number, col: number): string =>
	formatAddress({ row, col }).replace(/^([A-Z]+)(\d+)$/, '$$$1$$$2');

/** `=Sheet1!$A$1:$B$2` for the current selection. */
export function selectionReference(ctx: EditorContext): string {
	const ws = ctx.workbook()?.sheets[ctx.activeSheet()];
	const r = ctx.selection.get().ranges[0];
	if (!ws || !r) return '';
	const a = absolute(r.start.row, r.start.col);
	const b = absolute(r.end.row, r.end.col);
	return `=${quoteSheetName(ws.name)}!${a === b ? a : `${a}:${b}`}`;
}

export function openDefineName(
	ctx: EditorContext,
	props: DefineNameProps = {},
): Promise<DefinedName | undefined> {
	const session = ctx.session();
	if (!session) return Promise.resolve(undefined);
	const workbook = session.workbook;
	const existing = props.name
		? workbook.definedNames.find((n) => n.name === props.name && n.localSheet === props.localSheet)
		: undefined;
	const name = textInput(ctx, existing?.name ?? '');
	const scope = select(
		ctx,
		[
			['workbook', ctx.t('Workbook')],
			...workbook.sheets.map((s, i) => [String(i), s.name] as const),
		],
		existing?.localSheet === undefined ? 'workbook' : String(existing.localSheet),
		true,
	);
	if (existing) scope.disabled = true;
	const comment = textArea(ctx, existing?.comment ?? '', 3);
	const refersTo = textInput(ctx, existing ? `=${existing.formula}` : selectionReference(ctx));
	return showDialog<DefinedName>(ctx, {
		name: 'define-name',
		heading: existing ? 'Edit Name' : 'New Name',
		body: [
			field(ctx, 'Name:', name),
			field(ctx, 'Scope:', scope),
			field(ctx, 'Comment:', comment),
			field(ctx, 'Refers to:', refersTo),
		],
		opened: () => name.focus(),
		submit: () => {
			const value = name.value.trim();
			const problem = validateDefinedName(value);
			if (problem) return invalid(ctx, name, problem);
			const localSheet = scope.value === 'workbook' ? undefined : Number(scope.value);
			const clash = workbook.definedNames.some(
				(n) =>
					n !== existing &&
					n.name.toLowerCase() === value.toLowerCase() &&
					n.localSheet === localSheet,
			);
			if (clash)
				return invalid(ctx, name, 'The name that you entered already exists. Enter a unique name.');
			const formula = refersTo.value.trim().replace(/^=/, '');
			if (!formula) return invalid(ctx, refersTo, 'Enter a reference or a formula.');
			const next: DefinedName = { name: value, formula };
			if (localSheet !== undefined) next.localSheet = localSheet;
			if (comment.value.trim()) next.comment = comment.value.trim();
			session.batch(existing ? 'Edit name' : 'Define name', () => {
				if (existing && existing.name.toLowerCase() !== value.toLowerCase())
					session.deleteDefinedName(existing.name, existing.localSheet);
				session.setDefinedName(next);
			});
			return next;
		},
	});
}
