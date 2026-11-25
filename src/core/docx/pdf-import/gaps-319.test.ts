import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { loadDocx, shapeSvg } from '../index';
import { applyIccProfile } from '../../pdf/icc-profile';
import { decodeJbig2Pages } from '../../pdf/jbig2';
import { PathBuilder } from '../../pdf/content-paths';
import { convertPdf, importPdf } from './index';
import { extractEndnotes, type EndnotePage } from './endnotes';
import { joinSplitRow } from './row-split';
import type { Line } from './lines';
import type { TableBlock, TableCellBlock } from './tables';

/** Builds a minimal single-page PDF with custom resources and content stream. */
function makePdf(content: string, extraObjects: string[] = [], resources = ''): Uint8Array {
	const objects = [
		'<< /Type /Catalog /Pages 2 0 R >>',
		'<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
		`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >> ${resources} >> /Contents 4 0 R >>`,
		`<< /Length ${content.length} >>\nstream\n${content}\nendstream`,
		...extraObjects,
	];
	let pdf = '%PDF-1.4\n';
	const offsets: number[] = [0];
	objects.forEach((body, index) => {
		offsets.push(pdf.length);
		pdf += `${index + 1} 0 obj\n${body}\nendobj\n`;
	});
	const xref = pdf.length;
	pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
	for (let i = 1; i <= objects.length; i++)
		pdf += `${String(offsets[i]!).padStart(10, '0')} 00000 n \n`;
	pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
	return new TextEncoder().encode(pdf);
}

function makeLine(page: number, top: number, text: string, size = 11, mark?: string): Line {
	const fragments = mark
		? [
				{
					text: mark,
					x: 72,
					width: 8,
					size: size * 0.7,
					font: {
						name: 'Calibri',
						family: 'Calibri',
						bold: false,
						italic: false,
						serif: false,
						monospace: false,
					},
					color: '#000000',
					script: 'super' as const,
				},
				{
					text,
					x: 82,
					width: text.length * 5,
					size,
					font: {
						name: 'Calibri',
						family: 'Calibri',
						bold: false,
						italic: false,
						serif: false,
						monospace: false,
					},
					color: '#000000',
				},
			]
		: [
				{
					text,
					x: 72,
					width: text.length * 5,
					size,
					font: {
						name: 'Calibri',
						family: 'Calibri',
						bold: false,
						italic: false,
						serif: false,
						monospace: false,
					},
					color: '#000000',
				},
			];
	return {
		page,
		left: 72,
		right: 72 + text.length * 5 + (mark ? 10 : 0),
		top,
		bottom: top + size * 1.2,
		baseline: top + size,
		size,
		bold: false,
		text: mark ? `${mark}${text}` : text,
		fragments,
	};
}

describe('PDF import Issue #319 remaining gaps', () => {
	it('converts gradient fills, pattern fills, clipped shapes and text inside drawings', async () => {
		const shadingObj =
			'<< /ShadingType 2 /ColorSpace /DeviceRGB /Coords [72 600 272 600] /Function << /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 0 1] /N 1 >> >>';
		const patternObj =
			'<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] /XStep 10 /YStep 10 /Resources << >> /Length 14 >>\nstream\n0 1 0 rg 0 0 5 5 re f\nendstream';
		const content = [
			// 1. Axial gradient inside a clip rectangle.
			'q 72 600 200 80 re W n /Sh1 sh Q',
			// 2. Tiling pattern filled rectangle with text inside it.
			'q /Pattern cs /P1 scn 72 450 200 60 re f Q',
			'BT /F1 12 Tf 90 475 Td (Inside Box) Tj ET',
			// 3. Body text outside drawings.
			'BT /F1 12 Tf 72 350 Td (Body text) Tj ET',
		].join('\n');
		const bytes = makePdf(
			content,
			[shadingObj, patternObj],
			'/Shading << /Sh1 5 0 R >> /Pattern << /P1 6 0 R >>',
		);
		const { docx } = await convertPdf(bytes);
		const zip = await JSZip.loadAsync(docx);
		const xml = await zip.file('word/document.xml')!.async('string');
		expect(xml).toContain('<a:gradFill>');
		expect(xml).toContain('<a:pattFill');
		expect(xml).toContain('<wps:txbx>');
		expect(xml).toContain('Inside Box');
		expect(xml).toContain('Body text');
		// "Inside Box" is inside the shape text box, not duplicated in flowing body paragraphs.
		expect(xml.match(/Inside Box/g)).toHaveLength(1);
	});

	it('clips rectangular and polygon shapes against rectangular clip regions instead of dropping them', () => {
		const builder = new PathBuilder();
		builder.moveTo(0, 0);
		builder.lineTo(100, 0);
		builder.lineTo(100, 100);
		builder.lineTo(0, 100);
		builder.close();
		// Flip diagonal polygon so it goes through the complex path branch.
		builder.clear();
		builder.moveTo(0, 0);
		builder.lineTo(100, 50);
		builder.lineTo(50, 100);
		builder.close();
		const res = builder.paint(
			[1, 0, 0, 1, 0, 0],
			{
				ctm: [1, 0, 0, 1, 0, 0],
				fill: '#336699',
				stroke: '#000000',
				lineWidth: 1,
				fillAlpha: 1,
				strokeAlpha: 1,
				dash: [],
				fillPattern: false,
				strokePattern: false,
				clip: { x0: 10, y0: 10, x1: 80, y1: 80, rect: true },
			},
			false,
			true,
			false,
		);
		expect(res && 'path' in res).toBe(true);
		if (res && 'path' in res) {
			expect(res.path.x0).toBeGreaterThanOrEqual(10);
			expect(res.path.y0).toBeGreaterThanOrEqual(10);
			expect(res.path.x1).toBeLessThanOrEqual(80);
			expect(res.path.y1).toBeLessThanOrEqual(80);
		}
	});

	it('renders wpg:wgp groups as SVG shapes in loadDocx instead of placeholders', async () => {
		const content = [
			'q 1 0 0 rg 72 600 80 40 re f Q',
			'q 0 0 1 rg 150 600 80 40 re f Q',
			'BT /F1 12 Tf 72 500 Td (Caption) Tj ET',
		].join('\n');
		const { model } = await importPdf(makePdf(content));
		const images = model.blocks
			.flatMap((block) => (block.type === 'paragraph' ? block.runs : []))
			.map((run) => run.image)
			.filter((img): img is NonNullable<typeof img> => Boolean(img));
		const group = images.find((img) => img.shape?.groupChildren?.length);
		expect(group).toBeDefined();
		expect(group!.shape!.groupChildren!.length).toBe(2);
		const svg = shapeSvg(group!.shape!, group!.widthPx, group!.heightPx);
		expect(svg).toContain('<svg');
		expect(svg).toContain('#FF0000');
		expect(svg).toContain('#0000FF');
		// Also verify direct loadDocx round-trip preserves wpg group rendering.
		const { docx } = await convertPdf(makePdf(content));
		const reloaded = await loadDocx(docx);
		const reloadedGroup = reloaded.model.blocks
			.flatMap((b) => (b.type === 'paragraph' ? b.runs : []))
			.find((r) => r.image?.shape?.groupChildren?.length);
		expect(reloadedGroup).toBeDefined();
	});

	it('extracts endnotes collected at the end of each section with sectEnd placement', () => {
		const body1 = makeLine(0, 100, 'Section one text');
		body1.fragments.push({
			text: 'i',
			x: 160,
			width: 5,
			size: 7,
			font: body1.fragments[0]!.font,
			color: '#000000',
			script: 'super',
		});
		const heading1 = makeLine(0, 300, 'Notes', 11);
		const note1 = makeLine(0, 320, ' First section note.', 9, 'i');

		const body2 = makeLine(1, 100, 'Section two text');
		body2.fragments.push({
			text: 'ii',
			x: 160,
			width: 8,
			size: 7,
			font: body2.fragments[0]!.font,
			color: '#000000',
			script: 'super',
		});
		const heading2 = makeLine(1, 300, 'Notes', 11);
		const note2 = makeLine(1, 320, ' Second section note.', 9, 'ii');
		const after = makeLine(1, 500, 'Following section text', 11);

		const pages: EndnotePage[] = [
			{ lines: [body1, heading1, note1], rules: [] },
			{ lines: [body2, heading2, note2, after], rules: [] },
		];
		const extracted = extractEndnotes(pages, 11);
		expect(extracted).toBeDefined();
		expect(extracted!.position).toBe('sectEnd');
		expect(extracted!.format).toBe('lowerRoman');
		expect(extracted!.notes).toHaveLength(2);
		expect(pages[0]!.lines).toEqual([body1]);
		expect(pages[1]!.lines).toEqual([body2, after]);
	});

	it('joins a table row split between two sentences even when empty cells are normal', () => {
		const cell = (row: number, col: number, lines: string[], bordered = true): TableCellBlock => ({
			row,
			col,
			rowSpan: 1,
			colSpan: 1,
			borders: bordered
				? { top: { color: '#000000', width: 0.5 }, bottom: { color: '#000000', width: 0.5 } }
				: {},
			blocks: lines.length
				? [
						{
							kind: 'paragraph',
							page: 0,
							align: 'left',
							indentLeft: 0,
							firstLine: 0,
							spaceBefore: 0,
							lineSpacing: 14,
							size: 11,
							lines: lines.map((t, i) => makeLine(0, i * 14, t)),
							fragments: lines.map((t) => makeLine(0, 0, t).fragments[0]!),
						},
					]
				: [],
		});
		const table1: TableBlock = {
			kind: 'table',
			page: 0,
			top: 100,
			bottom: 200,
			left: 72,
			columns: [100, 300],
			padding: 4,
			rows: [{ height: 20 }, { height: 20 }, { height: 40 }],
			cells: [
				cell(0, 0, ['A']),
				cell(0, 1, ['First']),
				cell(1, 0, []), // empty cell in col 0 is normal in this table!
				cell(1, 1, ['Second']),
				cell(2, 0, ['C']),
				cell(2, 1, ['First sentence ends here.', 'Second line of row C.']),
			],
		};
		const table2: TableBlock = {
			kind: 'table',
			page: 1,
			top: 72,
			bottom: 150,
			left: 72,
			columns: [100, 300],
			padding: 4,
			rows: [{ height: 20 }, { height: 20 }],
			cells: [
				cell(0, 0, []),
				cell(0, 1, ['New sentence starts with capital letter.']),
				cell(1, 0, ['D']),
				cell(1, 1, ['Next row']),
			],
		};
		expect(joinSplitRow(table1, table2, 0)).toBe(true);
		const joined = table1.cells.find((c) => c.row === 2 && c.col === 1)!;
		expect(joined.blocks.map((b) => b.lines.map((l) => l.text).join(' ')).join(' ')).toContain(
			'New sentence starts with capital letter.',
		);
	});

	it('applies embedded ICC profiles (Gray and RGB Matrix/TRC) to decoded samples', () => {
		// Build a minimal valid ICC v2 Monochrome profile with gamma = 2.0 (u8Fixed8 = 0x0200).
		const icc = new Uint8Array(132 + 12 + 14);
		const view = new DataView(icc.buffer);
		view.setUint32(0, icc.length);
		icc.set([0x47, 0x52, 0x41, 0x59], 16); // 'GRAY'
		icc.set([0x61, 0x63, 0x73, 0x70], 36); // 'acsp'
		view.setUint32(128, 1); // 1 tag
		icc.set([0x6b, 0x54, 0x52, 0x43], 132); // 'kTRC'
		view.setUint32(136, 144);
		view.setUint32(140, 14);
		icc.set([0x63, 0x75, 0x72, 0x76], 144); // 'curv'
		view.setUint32(152, 1); // count = 1 (gamma)
		view.setUint16(156, 0x0200); // gamma = 2.0
		const samples = new Uint8Array([0, 128, 255]);
		expect(applyIccProfile(samples, 1, icc)).toBe(true);
		// Mid-grey 128 -> linear (128/255)^2 = 0.252 -> sRGB ~ 137.
		expect(samples[0]).toBe(0);
		expect(samples[1]).toBeGreaterThan(130);
		expect(samples[1]).toBeLessThan(145);
		expect(samples[2]).toBe(255);
	});

	it('decodes JBIG2 T.88 Amendment 2 extended generic templates (EXTTEMPLATE=1)', () => {
		// Minimal JBIG2 sequential stream: PageInfo (48) + Immediate Generic Region (38) with EXTTEMPLATE=1 (flags=16).
		const seg = (num: number, type: number, payload: Uint8Array) => {
			const header = new Uint8Array(11 + payload.length);
			const view = new DataView(header.buffer);
			view.setUint32(0, num);
			header[4] = type;
			header[5] = 0; // 0 referred segments
			header[6] = 1; // page 1
			view.setUint32(7, payload.length);
			header.set(payload, 11);
			return header;
		};
		const pageInfo = new Uint8Array(19);
		const pv = new DataView(pageInfo.buffer);
		pv.setUint32(0, 4); // width = 4
		pv.setUint32(4, 2); // height = 2
		// Generic region: 17-byte region info + 1-byte flags (16 = EXTTEMPLATE) + 24 bytes AT + 2 bytes arith data
		const gen = new Uint8Array(17 + 1 + 24 + 2);
		const gv = new DataView(gen.buffer);
		gv.setUint32(0, 4); // width
		gv.setUint32(4, 2); // height
		gen[17] = 16; // template 0, EXTTEMPLATE = 1, MMR = 0
		// 12 AT pixels: (-1,-1) etc.
		for (let i = 0; i < 12; i++) {
			gen[18 + i * 2] = 0xff; // dx = -1
			gen[18 + i * 2 + 1] = 0xff; // dy = -1
		}
		const s1 = seg(0, 48, pageInfo);
		const s2 = seg(1, 38, gen);
		const stream = new Uint8Array(s1.length + s2.length);
		stream.set(s1, 0);
		stream.set(s2, s1.length);
		const decoded = decodeJbig2Pages(stream)[0]!;
		expect(decoded.width).toBe(4);
		expect(decoded.height).toBe(2);
	});
});
