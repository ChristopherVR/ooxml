// Create Table (Insert > Table, Format as Table, Ctrl+T): the data range and the header flag.
import { type Table, formatRange, parseRange } from '@christophervr/xlsx-core';
import { guessHeader, regionOf, target } from '../commands/util.js';
import type { EditorContext } from '../context.js';
import { checkbox, field, textInput } from './fields.js';
import { showDialog } from './frame.js';

export interface CreateTableProps {
	styleName?: string;
}

const cleanRef = (value: string): string =>
	value.trim().replace(/^=/, '').replace(/^.*!/, '').replace(/\$/g, '');

export function openCreateTable(
	ctx: EditorContext,
	props: CreateTableProps = {},
): Promise<Table | undefined> {
	const t = target(ctx);
	if (!t) return Promise.resolve(undefined);
	const region = regionOf(t);
	const ref = textInput(ctx, `=${formatRange(region).replace(/([A-Z]+)(\d+)/g, '$$$1$$$2')}`);
	const headers = checkbox(ctx, 'My table has headers', guessHeader(t.ws, region));
	return showDialog<Table>(ctx, {
		name: 'create-table',
		heading: 'Create Table',
		body: [field(ctx, 'Where is the data for your table?', ref), headers.wrapper],
		opened: () => {
			ref.focus();
			ref.select();
		},
		submit: () => {
			const range = parseRange(cleanRef(ref.value));
			if (!range) {
				ctx.toast(ctx.t('The reference is not valid.'), 'warning');
				ref.focus();
				return undefined;
			}
			try {
				return t.session.createTable(t.sheet, range, headers.input.checked, props.styleName);
			} catch (error) {
				ctx.toast(ctx.t(error instanceof Error ? error.message : String(error)), 'error');
				return undefined;
			}
		},
	});
}
