/**
 * What the chapters of the report share: the document under construction, the outline as it
 * grows, and the cross-reference, caption and citation helpers that keep stored field results in
 * step with what they point at.
 */
import type { Paragraph, Table, TextRun } from '../../../../src/core/docx/index';
import { Doc, rich, type Inline, type Runs } from '../kit/builder';
import { equation } from '../kit/drawing';
import { caption, field, pageRef, ref } from '../kit/fields';
import { table, tw, type TableSpec } from '../kit/tables';
import { PAGES } from './pages';
import { TEXT_WIDTH } from './styles';

export interface OutlineEntry {
	level: 1 | 2 | 3;
	key: string;
	/** `2.1`, or empty for an unnumbered heading. */
	number: string;
	text: string;
}

/** The figures and tables of the report, in order: their numbers are their positions. */
export const FIGURES = ['site', 'flow', 'energy', 'curve'] as const;
export const TABLES = [
	'spec',
	'instruments',
	'quarterly',
	'environment',
	'monthly',
	'cost',
] as const;
export type FigureKey = (typeof FIGURES)[number];
export type TableKey = (typeof TABLES)[number];

export interface CaptionEntry {
	label: 'Figure' | 'Table';
	number: number;
	key: string;
	text: string;
}

/** Section numbers by key from a previous pass, so text can refer to a later section. */
export type KnownSections = ReadonlyMap<string, string>;

export const tocMark = (kind: 'sec' | 'fig' | 'tab', key: string) => `_Toc_${kind}_${key}`;
export const refMark = (kind: 'sec' | 'fig' | 'tab' | 'eq', key: string) => `_Ref_${kind}_${key}`;

/** The page Word shows for a bookmark, as measured in Word (`pages.ts`). */
export const pageOf = (mark: string): string => PAGES[mark] ?? '0';

export class Report {
	readonly doc = new Doc();
	readonly outline: OutlineEntry[] = [];
	readonly captions: CaptionEntry[] = [];
	private readonly counters = [0, 0, 0];
	private equations = 0;
	private readonly equationNumbers = new Map<string, number>();

	constructor(private readonly known: KnownSections) {}

	/** A numbered heading. Its `_Toc` and `_Ref` bookmarks bracket the paragraph. */
	h(level: 1 | 2 | 3, key: string, text: string): Paragraph {
		this.counters[level - 1]! += 1;
		for (let deeper = level; deeper < 3; deeper++) this.counters[deeper] = 0;
		const number = this.counters.slice(0, level).join('.');
		this.outline.push({ level, key, number, text });
		return this.doc.p(`Heading${level}`, text, {
			bookmarks: [tocMark('sec', key), refMark('sec', key)],
		});
	}

	/** A level-one heading without a number that the table of contents still lists. */
	plainHeading(key: string, text: string, props: Partial<Paragraph> = {}): Paragraph {
		this.outline.push({ level: 1, key, number: '', text });
		return this.doc.p('HeadingUnnumbered', text, {
			bookmarks: [tocMark('sec', key), refMark('sec', key)],
			...props,
		});
	}

	/** Body text from markup (see `rich`). */
	p(
		markup: string,
		inserts: Record<string, Inline> = {},
		props: Partial<Paragraph> = {},
	): Paragraph {
		return this.doc.p('Normal', rich(markup, inserts), props);
	}

	bullets(items: readonly (string | Runs)[]): Paragraph[] {
		return items.map((item) =>
			this.doc.p('ListBullet', typeof item === 'string' ? rich(item) : item),
		);
	}

	/** "Section 3.2" as a cross-reference to the heading's paragraph number. */
	sec(key: string): Runs {
		return [
			{ text: 'Section ' },
			...ref(refMark('sec', key), this.known.get(key) ?? '0', '\\r \\h'),
		];
	}

	fig(key: FigureKey): Runs {
		return ref(refMark('fig', key), `Figure ${FIGURES.indexOf(key) + 1}`, '\\h', 'CaptionLabel');
	}

	tab(key: TableKey): Runs {
		return ref(refMark('tab', key), `Table ${TABLES.indexOf(key) + 1}`, '\\h', 'CaptionLabel');
	}

	/** The page a figure, table or section is on, as a PAGEREF. */
	page(kind: 'sec' | 'fig' | 'tab', key: string): Runs {
		const mark = kind === 'sec' ? refMark('sec', key) : tocMark(kind, key);
		return pageRef(mark, pageOf(mark));
	}

	figureCaption(key: FigureKey, text: string, inserts: Record<string, Inline> = {}): Paragraph {
		const number = FIGURES.indexOf(key) + 1;
		this.captions.push({ label: 'Figure', number, key, text: plain(text) });
		return caption(this.doc, 'Figure', number, rich(text, inserts), {
			ref: refMark('fig', key),
			toc: tocMark('fig', key),
		});
	}

	/** A table caption sits above its table and stays with it. */
	tableCaption(key: TableKey, text: string): Paragraph {
		const number = TABLES.indexOf(key) + 1;
		this.captions.push({ label: 'Table', number, key, text: plain(text) });
		return caption(
			this.doc,
			'Table',
			number,
			rich(text),
			{ ref: refMark('tab', key), toc: tocMark('tab', key) },
			{ keepNext: true, spacingBeforeTwips: tw(200), spacingAfterTwips: tw(100) },
		);
	}

	/** A data table in the report's table style. */
	dataTable(spec: Omit<TableSpec, 'style' | 'textStyle'>): Table {
		return table(this.doc, {
			style: 'SaltmarshTable',
			textStyle: 'TableText',
			look: {
				firstRow: true,
				lastRow: false,
				firstColumn: false,
				lastColumn: false,
				noHBand: false,
				noVBand: true,
			},
			cantSplit: true,
			keepTogether: true,
			...spec,
		});
	}

	/**
	 * A numbered display equation: a borderless three-column grid holds the centred equation and
	 * its number, the usual Word layout for a number at the right margin.
	 */
	equation(key: string, latex: string, width = TEXT_WIDTH): Table {
		const number = ++this.equations;
		this.equationNumbers.set(key, number);
		const side = 850;
		return table(this.doc, {
			style: 'LayoutGrid',
			grid: [side, width - 2 * side, side],
			headerRows: 0,
			look: {
				firstRow: false,
				lastRow: false,
				firstColumn: false,
				lastColumn: false,
				noHBand: true,
				noVBand: true,
			},
			rows: [
				[
					{ content: '', style: 'EquationNumber' },
					{ content: [equation(latex, true)], style: 'Equation', vAlign: 'center' },
					{
						content: [
							...this.doc.bookmark(refMark('eq', key), [
								{ text: '(' },
								...field('SEQ Equation \\* ARABIC', String(number)),
								{ text: ')' },
							]),
						],
						style: 'EquationNumber',
						vAlign: 'center',
					},
				],
			],
		});
	}

	/** "Equation (2)" as a cross-reference; the equation must already be in the document. */
	eq(key: string): Runs {
		const number = this.equationNumbers.get(key);
		if (!number) throw new Error(`Equation ${key} is referenced before it is defined`);
		return [{ text: 'Equation ' }, ...ref(refMark('eq', key), `(${number})`)];
	}

	/**
	 * A citation: a CITATION field in a citation content control, as Word's References tab
	 * inserts it. `text` is what Word's APA style renders for the source.
	 */
	cite(tag: string, text: string): Runs {
		const control = {
			id: String(900000 + this.doc.nextMarkId()),
			level: 'inline' as const,
			kind: 'citation' as const,
		};
		return field(`CITATION ${tag} \\l 2057`, text).map((run): TextRun => ({
			...run,
			contentControl: control,
		}));
	}
}

/** Caption markup without its emphasis marks, for the lists of figures and tables. */
const plain = (markup: string): string => markup.replace(/\*\*|\*|\^|~|\{[^}]*\}/gu, '');
