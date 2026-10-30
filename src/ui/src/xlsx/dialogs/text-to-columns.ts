// Text to Columns: splits the first selected column into the columns starting at the destination,
// at delimiters or at fixed character positions. The splitting rules live in ooxml-core/xlsx.
import {
	type CellAddress,
	formatAddress,
	getCell,
	parseAddress,
	parseBreakPositions,
	splitDelimited,
	splitFixedWidth,
	type SplitOptions,
} from 'ooxml-core/xlsx';
import { target } from 'ooxml-core/xlsx/ui';
import type { EditorContext } from 'ooxml-core/xlsx/ui';
import {
	checkbox,
	el,
	field,
	fieldset,
	invalid,
	radios,
	row,
	select,
	text,
	textInput,
} from './fields';
import { showDialog } from './frame';

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
	const kind = radios(
		ctx,
		'Original data type',
		[
			['delimited', 'Delimited'],
			['fixed', 'Fixed width'],
		],
		'delimited',
	);
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
	const breaks = textInput(ctx);
	breaks.placeholder = '5, 12';
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
	/** The line splitter for the chosen type; undefined when the break positions are not valid. */
	const splitter = (): ((line: string) => string[]) | undefined => {
		if (kind.get() === 'fixed') {
			const positions = parseBreakPositions(breaks.value);
			return positions ? (line) => splitFixedWidth(line, positions) : undefined;
		}
		const opts = options();
		return (line) => splitDelimited(line, opts);
	};
	const refresh = (): void => {
		const split = splitter();
		breaks.setAttribute('aria-invalid', split ? 'false' : 'true');
		const table = el(ctx, 'table');
		if (split)
			for (const line of lines.slice(0, 5)) {
				const tr = el(ctx, 'tr');
				for (const part of split(line)) {
					const td = el(ctx, 'td');
					td.textContent = part;
					tr.append(td);
				}
				table.append(tr);
			}
		preview.replaceChildren(table);
	};
	const delimitedPart = el(ctx, 'div');
	delimitedPart.append(
		fieldset(
			ctx,
			'Delimiters',
			row(ctx, tab.wrapper, semicolon.wrapper, comma.wrapper, space.wrapper),
			row(ctx, other.wrapper, otherChar),
		),
		consecutive.wrapper,
		field(ctx, 'Text qualifier:', qualifier),
	);
	const fixedPart = el(ctx, 'div');
	fixedPart.append(
		field(ctx, 'Break positions:', breaks),
		text(ctx, 'Each number ends a column after that character, for example 5, 12.'),
	);
	fixedPart.hidden = true;
	const showKind = (): void => {
		const fixed = kind.get() === 'fixed';
		delimitedPart.hidden = fixed;
		fixedPart.hidden = !fixed;
		refresh();
	};
	kind.inputs.forEach((input) => input.addEventListener('change', showKind));
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
	breaks.addEventListener('input', refresh);
	refresh();
	return showDialog<number>(ctx, {
		name: 'text-to-columns',
		heading: 'Convert Text to Columns',
		wide: true,
		body: [
			kind.element,
			delimitedPart,
			fixedPart,
			field(ctx, 'Destination:', destination),
			preview,
		],
		opened: () => kind.inputs[0]?.focus(),
		submit: () => {
			const dest: CellAddress | undefined = parseAddress(
				destination.value.trim().replace(/^=/, '').replace(/^.*!/, ''),
			);
			if (!dest) {
				ctx.toast(ctx.t('The reference is not valid.'), 'warning');
				destination.focus();
				return undefined;
			}
			const split = splitter();
			if (!split)
				return invalid(
					ctx,
					breaks,
					'Enter break positions as whole numbers from 1, separated by commas.',
				);
			let written = 0;
			t.session.batch('Text to Columns', () => {
				lines.forEach((line, i) => {
					if (!line) return;
					split(line).forEach((part, j) => {
						t.session.setCellInput(t.sheet, dest.row + i, dest.col + j, part);
						written++;
					});
				});
			});
			return written;
		},
	});
}
