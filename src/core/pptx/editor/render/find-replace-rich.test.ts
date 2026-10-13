import { describe, expect, it } from 'vitest';
import type {
	PptxElement,
	PptxSlide,
	TablePptxElement,
	SmartArtPptxElement,
	TextPptxElement,
} from '../../index';
import { findInSlides, replaceInSlides } from './find-replace';

const slidesOf = (...elements: PptxElement[]): PptxSlide[] => [
	{ id: 's', rId: 'r', slideNumber: 1, elements },
];
const text = (): TextPptxElement => ({
	id: 'text',
	type: 'text',
	x: 0,
	y: 0,
	width: 200,
	height: 50,
	text: 'Annual report',
	textSegments: [
		{ text: 'Annual re', style: { bold: true } },
		{ text: 'port', style: { italic: true, color: '#123456' } },
	],
});

describe('rich and nested find/replace', () => {
	it('finds across formatting runs and preserves both formats and unmatched suffixes', () => {
		const source = slidesOf(text());
		expect(findInSlides(source, 'report')).toHaveLength(1);
		const result = replaceInSlides(source, 'report', 'summary');
		expect(result.replacements).toBe(1);
		const element = result.slides[0]!.elements[0] as TextPptxElement;
		expect(element.text).toBe('Annual summary');
		expect(element.textSegments).toEqual([
			{ text: 'Annual summary', style: { bold: true } },
			{ text: '', style: { italic: true, color: '#123456' } },
		]);
		expect((source[0]!.elements[0] as TextPptxElement).text).toBe('Annual report');
	});

	it('keeps paragraph boundaries and ignores overlapping replace-all matches', () => {
		const element = text();
		element.textSegments = [
			{ text: 'aaaa', style: {} },
			{ text: '', style: {}, isParagraphBreak: true },
			{ text: 'word', style: {} },
		];
		const source = slidesOf(element);
		expect(findInSlides(source, 'aaword')).toHaveLength(0);
		const result = replaceInSlides(source, 'aa', 'x');
		expect(result.replacements).toBe(2);
		expect((result.slides[0]!.elements[0] as TextPptxElement).text).toBe('xx\nword');
	});

	it('selects the enclosing group and edits its nested text without cloning its sibling', () => {
		const sibling = {
			...text(),
			id: 'sibling',
			text: 'Keep',
			textSegments: [{ text: 'Keep', style: {} }],
		};
		const group = {
			id: 'group',
			type: 'group',
			x: 0,
			y: 0,
			width: 200,
			height: 100,
			children: [
				{ id: 'nested', type: 'group', x: 0, y: 0, width: 200, height: 100, children: [text()] },
				sibling,
			],
		} as PptxElement;
		const source = slidesOf(group);
		expect(findInSlides(source, 'report')[0]).toMatchObject({
			elementId: 'group',
			elementPath: [0, 0],
		});
		const result = replaceInSlides(source, 'report', 'summary');
		const updated = result.slides[0]!.elements[0];
		if (updated?.type !== 'group') throw new Error('Expected group');
		expect(updated.children[1]).toBe(sibling);
		if (updated.children[0]?.type !== 'group') throw new Error('Expected nested group');
		expect((updated.children[0].children[0] as TextPptxElement).text).toBe('Annual summary');
	});

	it('finds plain text without a run model and retains raw offsets for Unicode case folding', () => {
		const element = text();
		delete element.textSegments;
		element.text = 'İ Annual report';
		const source = slidesOf(element);
		expect(findInSlides(source, 'annual')[0]!.startOffset).toBe(2);
		expect(
			(replaceInSlides(source, 'Annual', 'Final').slides[0]!.elements[0] as TextPptxElement).text,
		).toBe('İ Final report');
	});

	it('updates rich table cell XML without removing run properties or opaque content', () => {
		const rawXml = {
			'a:graphic': {
				'a:graphicData': {
					'a:tbl': {
						'a:tr': {
							'a:tc': {
								'a:txBody': {
									'a:bodyPr': { '@_anchor': 'ctr' },
									'a:p': {
										'a:pPr': { '@_algn': 'r' },
										'a:r': [
											{ 'a:rPr': { '@_b': '1' }, 'a:t': 'Annu' },
											{
												'a:rPr': { '@_i': '1', 'a:hlinkClick': { '@_r:id': 'rIdKeep' } },
												'a:t': 'al report',
											},
										],
										'a:endParaRPr': { '@_lang': 'fr-FR' },
									},
									'vendor:keep': { '@_value': 'opaque' },
								},
								'a:tcPr': { '@_anchor': 'ctr' },
							},
						},
					},
				},
			},
		};
		const table = {
			id: 'table',
			type: 'table',
			x: 0,
			y: 0,
			width: 200,
			height: 100,
			rawXml,
			tableData: {
				rows: [
					{
						cells: [
							{
								text: 'Annual report',
								textRuns: [
									{ text: 'Annu', bold: true },
									{ text: 'al report', italic: true },
								],
							},
						],
					},
				],
				columnWidths: [1],
			},
		} as TablePptxElement;
		const source = slidesOf(table);
		expect(findInSlides(source, 'Annual')[0]).toMatchObject({
			tableCell: { rowIndex: 0, colIndex: 0 },
		});
		const updated = replaceInSlides(source, 'Annual', 'Final').slides[0]!
			.elements[0] as TablePptxElement;
		expect(updated.tableData!.rows[0]!.cells[0]!.text).toBe('Final report');
		expect(updated.tableData!.rows[0]!.cells[0]!.textRuns).toEqual([
			{ text: 'Final', bold: true },
			{ text: ' report', italic: true },
		]);
		expect(JSON.stringify(updated.rawXml)).toContain('rIdKeep');
		expect(JSON.stringify(updated.rawXml)).toContain('vendor:keep');
		expect(JSON.stringify(rawXml)).toContain('al report');
	});

	it('updates nested SmartArt text and keeps its cached drawing in sync', () => {
		const diagram = {
			id: 'diagram',
			type: 'smartArt',
			x: 0,
			y: 0,
			width: 300,
			height: 200,
			smartArtData: {
				nodes: [{ id: 'root', text: 'Root', children: [{ id: 'child', text: 'Annual report' }] }],
				drawingShapes: [
					{ id: 'node-child', text: 'Annual report', x: 0, y: 0, width: 100, height: 30 },
				],
			},
		} as SmartArtPptxElement;
		const updated = replaceInSlides(slidesOf(diagram), 'report', 'summary').slides[0]!
			.elements[0] as SmartArtPptxElement;
		expect(updated.smartArtData!.nodes[0]!.children![0]!.text).toBe('Annual summary');
		expect(updated.smartArtData!.drawingShapes![0]!.text).toBe('Annual summary');
		expect(updated.smartArtData!.drawingDirty).toBe(true);
	});
});

describe('speaker-note search scope', () => {
	it('opts into rich notes and replaces across runs without changing slide text', () => {
		const source = slidesOf(text());
		source[0]!.notes = 'Annual report';
		source[0]!.notesSegments = [
			{ text: 'Annual re', style: { bold: true } },
			{ text: 'port', style: { italic: true } },
		];
		expect(findInSlides(source, 'report')).toHaveLength(1);
		const matches = findInSlides(source, 'report', { includeNotes: true });
		expect(matches).toHaveLength(2);
		expect(matches[1]).toMatchObject({
			notes: true,
			elementId: '',
			segmentIndex: 0,
			startOffset: 7,
		});
		const replaced = replaceInSlides(source, 'report', 'summary', { includeNotes: true });
		expect(replaced.replacements).toBe(2);
		expect(replaced.slides[0]!.notes).toBe('Annual summary');
		expect(replaced.slides[0]!.notesSegments).toEqual([
			{ text: 'Annual summary', style: { bold: true } },
			{ text: '', style: { italic: true } },
		]);
		expect(source[0]!.notes).toBe('Annual report');
	});
	it('uses edited plain notes instead of stale rich runs and preserves separate paragraphs', () => {
		const source = slidesOf();
		source[0]!.notes = 'report report';
		source[0]!.notesSegments = [{ text: 'Old text', style: { bold: true } }];
		const replaced = replaceInSlides(source, 'report', 'summary', { includeNotes: true });
		expect(replaced.replacements).toBe(2);
		expect(replaced.slides[0]!.notes).toBe('summary summary');
		source[0]!.notes = 'Annual\nreport';
		source[0]!.notesSegments = [
			{ text: 'Annual', style: {} },
			{ text: '\n', style: {}, isParagraphBreak: true },
			{ text: 'report', style: { italic: true } },
		];
		expect(findInSlides(source, 'Annualreport', { includeNotes: true })).toHaveLength(0);
		expect(
			replaceInSlides(source, 'report', 'summary', { includeNotes: true }).slides[0]!.notes,
		).toBe('Annual\nsummary');
	});
});
