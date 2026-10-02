// Text to Columns (delimited): splits the first selected column at the chosen delimiters into
// the columns starting at the destination. Fixed-width splitting is not offered.
import { type CellAddress, formatAddress, getCell, parseAddress } from '@christophervr/xlsx-core';
import { target } from '../commands/util.js';
import type { EditorContext } from '../context.js';
import { checkbox, el, field, fieldset, row, select, text, textInput } from './fields.js';
import { showDialog } from './frame.js';

export interface SplitOptions {
	delimiters: string[];
	consecutive: boolean;
	qualifier: string;
}

/** Splits one line at any of the delimiters, honouring a text qualifier (`"a,b",c`). */
export function splitLine(line: string, options: SplitOptions): string[] {
	const out: string[] = [];
	let current = '';
	let quoted = false;
	for (let i = 0; i < line.length; i++) {
		const ch = line[i] ?? '';
		if (options.qualifier && ch === options.qualifier) {
			if (quoted && line[i + 1] === options.qualifier) {
				current += ch;
				i++;
			} else quoted = !quoted;
			continue;
		}
		if (!quoted && options.delimiters.includes(ch)) {
			out.push(current);
			current = '';
			if (options.consecutive) while (options.delimiters.includes(line[i + 1] ?? '\u0000')) i++;
			continue;
		}
		current += ch;
	}
	out.push(current);
	return out;
}

const cellText = (value: unknown): string =>
	value === null || value === undefined
		? ''
		: typeof value === 'object'
			? String((value as { error: string }).error)
			: String(value);

export function openTextToColumns(ctx: EditorContext): Promise<number | undefined> {
	const t = target(ctx);
	if (!t) return Promise.resolve(undefined);
	const r = t.range;
	const col = r.start.col;
	const lines: string[] = [];
	for (let rowIndex = r.start.row; rowIndex <= r.end.row; rowIndex++)
		lines.push(cellText(getCell(t.ws, rowIndex, col)?.value));
	const tab = checkbox(ctx, 'Tab', true);
	const semicolon = checkbox(ctx, 'Semicolon');
	const comma = checkbox(ctx, 'Comma');
	const space = checkbox(ctx, 'Space');
	const other = checkbox(ctx, 'Other:');
	const otherChar = textInput(ctx);
	otherChar.maxLength = 1;
	otherChar.setAttribute('aria-label', ctx.t('Other delimiter'));
	const consecutive = checkbox(ctx, 'Treat consecutive delimiters as one');
	const qualifier = select(
		ctx,
		[
			['"', '"'],
			["'", "'"],
			['', ctx.t('(none)')],
		],
		'"',
		true,
	);
	const destination = textInput(ctx, formatAddress(r.start));
	const preview = el(ctx, 'div', 'xve-results');
	preview.setAttribute('aria-label', ctx.t('Data preview'));
	const options = (): SplitOptions => ({
		delimiters: [
			...(tab.input.checked ? ['\t'] : []),
			...(semicolon.input.checked ? [';'] : []),
			...(comma.input.checked ? [','] : []),
			...(space.input.checked ? [' '] : []),
			...(other.input.checked && otherChar.value ? [otherChar.value] : []),
		],
		consecutive: consecutive.input.checked,
		qualifier: qualifier.value,
	});
	const refresh = (): void => {
		const table = el(ctx, 'table');
		for (const line of lines.slice(0, 5)) {
			const tr = el(ctx, 'tr');
			for (const part of splitLine(line, options())) {
				const td = el(ctx, 'td');
				td.textContent = part;
				tr.append(td);
			}
			table.append(tr);
		}
		preview.replaceChildren(table);
	};
	for (const control of [
		tab.input,
		semicolon.input,
		comma.input,
		space.input,
		other.input,
		consecutive.input,
		qualifier,
	])
		control.addEventListener('change', refresh);
	otherChar.addEventListener('input', () => {
		if (otherChar.value) other.input.checked = true;
		refresh();
	});
	refresh();
	return showDialog<number>(ctx, {
		name: 'text-to-columns',
		heading: 'Convert Text to Columns',
		wide: true,
		body: [
			fieldset(
				ctx,
				'Delimiters',
				row(ctx, tab.wrapper, semicolon.wrapper, comma.wrapper, space.wrapper),
				row(ctx, other.wrapper, otherChar),
			),
			consecutive.wrapper,
			field(ctx, 'Text qualifier:', qualifier),
			field(ctx, 'Destination:', destination),
			preview,
			text(ctx, 'Only delimited text is supported; fixed width splitting is not available yet.'),
		],
		opened: () => tab.input.focus(),
		submit: () => {
			const dest: CellAddress | undefined = parseAddress(
				destination.value.trim().replace(/^=/, '').replace(/^.*!/, ''),
			);
			if (!dest) {
				ctx.toast(ctx.t('The reference is not valid.'), 'warning');
				destination.focus();
				return undefined;
			}
			const opts = options();
			let written = 0;
			t.session.batch('Text to Columns', () => {
				lines.forEach((line, i) => {
					if (!line) return;
					splitLine(line, opts).forEach((part, j) => {
						t.session.setCellInput(t.sheet, dest.row + i, dest.col + j, part);
						written++;
					});
				});
			});
			return written;
		},
	});
}
