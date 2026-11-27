/**
 * The form's building blocks: the content controls (core model objects with a title, a tag and
 * placeholder text) and the grid of labelled entries they sit in.
 */
import type {
	ContentControl,
	Paragraph,
	Table,
	TableBorders,
	TextRun,
} from '../../../../src/core/docx/index';
import { FONTS } from '../../lib/fiction';
import { t, W_NS, type Doc, type Runs } from '../kit/builder';
import { line, NO_LINE, table, type CellInput, type CellSpec } from '../kit/tables';
import { INK, TEXT_WIDTH } from './styles';

const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';

/** Hands out the ids of the form's controls: fixed, unique and in document order. */
export class Controls {
	private next = 410000;

	private control(
		kind: ContentControl['kind'],
		alias: string,
		tag: string,
		extra: Partial<ContentControl> = {},
	): ContentControl {
		return { id: String(++this.next), level: 'inline', kind, alias, tag, ...extra };
	}

	/** A control showing `placeholder` until it is filled. */
	private empty(control: ContentControl, placeholder: string): Runs {
		return [
			{
				text: placeholder,
				style: 'PlaceholderText',
				contentControl: { ...control, showingPlaceholder: true },
			},
		];
	}

	text(alias: string, tag: string, placeholder: string, multiLine = false): Runs {
		return this.empty(
			this.control('text', alias, tag, multiLine ? { multiLine } : {}),
			placeholder,
		);
	}

	/** A plain text control that already holds `value`. */
	filled(alias: string, tag: string, value: string, props: Partial<TextRun> = {}): Runs {
		return [{ ...props, text: value, contentControl: this.control('text', alias, tag) }];
	}

	choice(
		kind: 'dropDownList' | 'comboBox',
		alias: string,
		tag: string,
		options: readonly string[],
		placeholder: string,
	): Runs {
		const items = options.map((option) => ({ display: option, value: option }));
		return this.empty(this.control(kind, alias, tag, { items }), placeholder);
	}

	date(alias: string, tag: string, placeholder: string): Runs {
		return this.empty(
			this.control('date', alias, tag, { dateFormat: 'd MMMM yyyy', dateLocale: 'en-GB' }),
			placeholder,
		);
	}

	/**
	 * A check box and its label. The glyphs are the usual ballot boxes, drawn in the body font:
	 * where that font has no such glyph, the application falls back as it would for any text.
	 */
	check(alias: string, tag: string, label: string, checked = false): Runs {
		return [
			{
				text: checked ? '☒' : '☐',
				fontFamily: FONTS.body,
				contentControl: this.control('checkbox', alias, tag, {
					checked,
					checkedGlyph: { code: '2612', font: FONTS.body },
					uncheckedGlyph: { code: '2610', font: FONTS.body },
				}),
			},
			t(` ${label}   `),
		];
	}

	/** A block-level control around paragraphs: every paragraph that carries it is inside it. */
	block(
		kind: 'richText' | 'group',
		alias: string,
		tag: string,
		extra: Partial<ContentControl> = {},
	): ContentControl {
		return { ...this.control(kind, alias, tag, extra), level: 'block' };
	}

	/** A control whose properties the model has no field for, given as `w:sdtPr` children. */
	verbatim(
		level: ContentControl['level'],
		kind: ContentControl['kind'],
		alias: string,
		tag: string,
		typeXml: string,
		extraXml = '',
	): ContentControl {
		const control = this.control(kind, alias, tag);
		return {
			...control,
			level,
			sourceXml: `<w:sdtPr xmlns:w="${W_NS}" xmlns:w15="${W15_NS}"><w:alias w:val="${alias}"/><w:tag w:val="${tag}"/><w:id w:val="${control.id}"/>${extraXml}${typeXml}</w:sdtPr>`,
		};
	}

	picture(alias: string, tag: string): ContentControl {
		return this.control('picture', alias, tag);
	}
}

/** "Part A   Client details": a part heading with its letter in the accent colour. */
export const partHeading = (
	doc: Doc,
	letter: string,
	title: string,
	props: Partial<Paragraph> = {},
) => doc.p('Heading1', [t(`Part ${letter}`, { style: 'PartLabel' }), t(`   ${title}`)], props);

export const NO_BORDERS: TableBorders = {
	top: NO_LINE,
	left: NO_LINE,
	bottom: NO_LINE,
	right: NO_LINE,
	insideH: NO_LINE,
	insideV: NO_LINE,
};
const ENTRY_LINE: TableBorders = { bottom: line(INK.slate, 0.75), top: line(INK.white, 3) };

export const label = (text: string): CellSpec => ({
	content: text,
	style: 'FieldLabel',
	vAlign: 'center',
});
/** The cell a value is entered in: a paper-coloured panel on a rule. */
export const entry = (content: Runs, extra: Partial<CellSpec> = {}): CellSpec => ({
	content,
	style: 'FieldValue',
	fill: `#${INK.paper}`,
	borders: ENTRY_LINE,
	vAlign: 'center',
	...extra,
});

const LABEL_WIDTH = 1480;

/** A grid of labelled entries, two to a row. */
export function formGrid(doc: Doc, rows: CellInput[][]): Table {
	const value = Math.floor((TEXT_WIDTH - 2 * LABEL_WIDTH) / 2);
	return table(doc, {
		style: 'FormGrid',
		grid: [LABEL_WIDTH, value, LABEL_WIDTH, TEXT_WIDTH - 2 * LABEL_WIDTH - value],
		headerRows: 0,
		borders: NO_BORDERS,
		look: {
			firstRow: false,
			lastRow: false,
			firstColumn: false,
			lastColumn: false,
			noHBand: true,
			noVBand: true,
		},
		cantSplit: true,
		rows,
	});
}
