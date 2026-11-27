/** A shape's text: runs with their own character formats, text fields, and hyperlink rows. */
import { xmlEscape } from '../../lib/package';
import { type TextStyle, characterCells } from './styles';
import { cell, row, section } from './xml';

/** A stretch of text with its own character format, or a field showing a formula's value. */
export interface Run {
	text: string;
	/** The run's complete character format: font, size and colour at least. */
	style: TextStyle;
	/** A ShapeSheet formula (`Prop.Owner`); `text` is then its cached value. */
	field?: string;
	/** A number or date picture (`0`, `{{d MMM yyyy}}`); the field shows text when absent. */
	picture?: string;
	/** For a date field: the date as `yyyy-mm-dd`, while `text` is how the picture shows it. */
	date?: string;
}
export type Rich = string | readonly Run[];

/**
 * A complete Character row. A row a shape adds beyond those of its master and style starts from
 * zeros (a FontScale of 0 squeezes the run to nothing), so every cell is written.
 */
function runCells(style: TextStyle): string[] {
	const set = characterCells(style);
	const has = (name: string) => set.some((text) => text.startsWith(`<Cell N="${name}"`));
	const defaults: [string, string | number][] = [
		['Style', 0],
		['Case', 0],
		['Pos', 0],
		['FontScale', 1],
		['DblUnderline', 0],
		['Overline', 0],
		['Strikethru', 0],
		['DoubleStrikethrough', 0],
		['Letterspace', 0],
		['ColorTrans', 0],
		['AsianFont', 0],
		['ComplexScriptFont', 0],
		['ComplexScriptSize', -1],
	];
	if (!has('Font') || !has('Color') || !has('Size'))
		throw new Error('A text run needs a font, a colour and a size.');
	return [...set, ...defaults.filter(([name]) => !has(name)).map(([n, v]) => cell(n, v))];
}

export const hasField = (rich: Rich | undefined): rich is readonly Run[] =>
	typeof rich !== 'string' && (rich ?? []).some((run) => run.field);

/** The sections and Text element of a shape's text. */
export function textBody(rich: Rich | undefined): string {
	if (rich === undefined) return '';
	if (typeof rich === 'string') return `<Text>${xmlEscape(rich)}</Text>`;
	const characters: string[] = [];
	const fields: string[] = [];
	let text = '';
	rich.forEach((run, index) => {
		// Every run gets a complete row: Visio draws text unpredictably when a shape's own
		// Character section starts at a row other than the first.
		characters.push(row({ IX: index }, runCells(run.style)));
		text += `<cp IX="${index}"/>`;
		if (run.field) {
			const at = fields.length;
			const picture = run.picture ?? '@';
			fields.push(
				row({ IX: at }, [
					run.date
						? cell('Value', `${run.date}T00:00:00`, run.field, 'DATE')
						: run.picture
							? cell('Value', Number(run.text), run.field)
							: cell('Value', run.text, run.field, 'STR'),
					// The picture as a quoted string, the form the library's own field writer uses.
					cell('Format', picture, `"${picture}"`, 'STR'),
					cell('Type', run.date ? 5 : run.picture ? 2 : 0),
					cell('UICat', 0),
					cell('UICod', 255),
					cell('UIFmt', 0),
					cell('Calendar', 0),
					cell('ObjectKind', 0),
				]),
			);
			text += `<fld IX="${at}">${xmlEscape(run.text)}</fld>`;
		} else text += xmlEscape(run.text);
	});
	return `${section('Character', characters)}${section('Field', fields)}<Text>${text}</Text>`;
}

/** The Hyperlink section of a shape with one link. */
export function hyperlinkSection(link: {
	address: string;
	subAddress: string;
	description: string;
}): string {
	return section('Hyperlink', [
		row({ N: 'Row_1' }, [
			cell('Description', link.description),
			cell('Address', link.address),
			cell('SubAddress', link.subAddress),
			cell('ExtraInfo', ''),
			cell('Frame', ''),
			cell('SortKey', ''),
			cell('NewWindow', 0),
			cell('Default', 0),
			cell('Invisible', 0),
		]),
	]);
}
