import { writeFile } from 'node:fs/promises';
import { inflateSync } from 'node:zlib';
import JSZip from 'jszip';
import { loadDocx } from 'ooxml-core/docx';
import { createFakeMeasurer } from 'ooxml-core/docx/layout';
import { describe, expect, it } from 'vitest';
import { exportDocxPdf } from './export-pdf';
import { paleColor } from './pdf-balloon-comments';
import {
	balloonFrame,
	layoutBalloons,
	paneGeometry,
	wrapBalloon,
	type Balloon,
} from './pdf-balloons';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const measurer = createFakeMeasurer();
const letter = { widthPx: 816, heightPx: 1056, marginRightPx: 96 };

function inspect(bytes: Uint8Array) {
	const raw = Buffer.from(bytes).toString('latin1');
	const streams: string[] = [];
	for (const match of raw.matchAll(/<<([^]*?)>>\nstream\n/g)) {
		const start = match.index + match[0].length;
		const length = Number(/\/Length (\d+)/.exec(match[1]!)?.[1]);
		const data = bytes.subarray(start, start + length);
		if (
			/\/Filter \/FlateDecode/.test(match[1]!) &&
			!/\/Subtype \/(Image|Form)|\/Length1|\/N 3/.test(match[1]!)
		)
			streams.push(inflateSync(data).toString('latin1'));
	}
	return { raw, content: streams.join('\n') };
}

const run = (text: string) => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;

/** A comment with a reply, a run and a paragraph formatting change, an insertion and a deletion. */
async function reviewed() {
	const zip = new JSZip();
	const body =
		`<w:p><w:pPr><w:jc w:val="center"/><w:pPrChange w:id="1" w:author="Ada"><w:pPr/></w:pPrChange></w:pPr>${run('Kept ')}<w:ins w:id="2" w:author="Ada"><w:r><w:t>added</w:t></w:r></w:ins><w:del w:id="3" w:author="Bob"><w:r><w:delText xml:space="preserve"> gone</w:delText></w:r></w:del><w:commentRangeStart w:id="7"/>${run(' noted')}<w:commentRangeEnd w:id="7"/></w:p>` +
		`<w:p><w:r><w:rPr><w:b/><w:rPrChange w:id="4" w:author="Bob"><w:rPr/></w:rPrChange></w:rPr><w:t>Bold now</w:t></w:r><w:hyperlink r:id="rL">${run(' link')}</w:hyperlink></w:p>`;
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>${body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/></w:sectPr></w:body></w:document>`,
	);
	zip.file(
		'word/comments.xml',
		`<w:comments xmlns:w="${W}"><w:comment w:id="7" w:author="Cy Young" w:initials="CY" w:date="2026-03-04T05:06:07Z"><w:p><w:r><w:t>Check this</w:t></w:r></w:p></w:comment><w:comment w:id="8" w:author="Di" w:date="2026-03-05T00:00:00Z"><w:p><w:r><w:t>Done</w:t></w:r></w:p></w:comment></w:comments>`,
	);
	zip.file(
		'word/commentsExtended.xml',
		`<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"/>`,
	);
	zip.file(
		'word/_rels/document.xml.rels',
		`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rC" Type="${R}/comments" Target="comments.xml"/><Relationship Id="rL" Type="${R}/hyperlink" Target="https://example.com/" TargetMode="External"/></Relationships>`,
	);
	const { model } = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	// The reply thread, as commentsExtended.xml links it.
	const reply = model.comments?.find((comment) => comment.id === '8');
	if (reply) reply.parentId = '7';
	return model;
}

describe('markup balloon layout', () => {
	it('shrinks the page to fit Word\'s 3.7" pane beside the text, centred on the sheet', () => {
		const frame = balloonFrame(letter);
		// Word prints a Letter page with 1" margins at 75.9% (measured from its own PDF export).
		expect(frame.scale).toBeCloseTo(816 / 1075.2, 4);
		expect(frame.offsetYPt).toBeCloseTo((792 - 792 * frame.scale) / 2, 3);
		const pane = paneGeometry(letter);
		expect(pane.left).toBe(732);
		expect(pane.right).toBeCloseTo(1075.2, 3);
		// About 174.6 pt wide on the sheet, as Word's balloons are.
		expect(pane.balloonWidth * frame.scale * 0.75).toBeCloseTo(174.6, 0);
	});

	it('stacks balloons near their anchors without overlap and keeps them on the page', () => {
		const balloon = (top: number, text: string): Balloon => ({
			kind: 'comment',
			color: '#1a73e8',
			paragraphs: [[{ text }]],
			anchorX: 100,
			anchorTop: top,
			anchorBottom: top + 20,
			marker: false,
		});
		const items = [balloon(200, 'one'), balloon(200, 'two '.repeat(40)), balloon(210, 'three')];
		const { fits } = layoutBalloons(items, letter, measurer);
		expect(fits).toBe(true);
		const boxes = items.map((item) => item.box!);
		expect(boxes[0]!.y).toBe(200);
		for (let index = 1; index < boxes.length; index++)
			expect(boxes[index]!.y).toBeGreaterThanOrEqual(
				boxes[index - 1]!.y + boxes[index - 1]!.height,
			);
		expect(boxes[1]!.height).toBeGreaterThan(boxes[0]!.height * 2);
		const low = [balloon(1040, 'a'), balloon(1050, 'b')];
		layoutBalloons(low, letter, measurer);
		expect(low[1]!.box!.y + low[1]!.box!.height).toBeLessThanOrEqual(letter.heightPx);
		expect(low[0]!.box!.y + low[0]!.box!.height).toBeLessThanOrEqual(low[1]!.box!.y);
		const crowded = Array.from({ length: 80 }, (_, index) => balloon(index, `item ${index}`));
		expect(layoutBalloons(crowded, letter, measurer).fits).toBe(false);
	});

	it('wraps balloon text by words and keeps forced line breaks', () => {
		const lines = wrapBalloon(
			[[{ text: 'Commented [A1]: ', bold: true }, { text: 'alpha beta gamma\ndelta' }]],
			120,
			measurer,
		);
		expect(lines.length).toBeGreaterThanOrEqual(3);
		for (const line of lines)
			expect(line.reduce((sum, piece) => sum + piece.width, 0)).toBeLessThanOrEqual(
				120 + measurer.widthOf(' ', { family: 'x', sizePx: 12 }),
			);
		expect(
			lines
				.at(-1)!
				.map((piece) => piece.text.trim())
				.join(''),
		).toBe('delta');
		expect(paleColor('#1a73e8')).toBe('#d1e3fa');
	});
});

describe('exportDocxPdf markup balloons', () => {
	it("shows comments and formatting in balloons by default, as Word's print markup does", async () => {
		const model = await reviewed();
		const result = await exportDocxPdf(model, { measurer, markup: 'all' });
		await dump('balloons', result.bytes);
		const pdf = inspect(result.bytes);
		const frame = balloonFrame(letter);
		expect(pdf.content).toMatch(/^0\.758929 0 0 0\.758929 0 95\.464286 cm\n/);
		// The comment balloon: its label, text, author and date, and the reply, as real text.
		expect(pdf.content).toContain('(Commented [CY1]: ) Tj');
		expect(pdf.content).toMatch(/\(Check (this)?\) Tj/);
		expect(pdf.content).toMatch(/\(Cy Young, 2026-03-04 05:06\) Tj/);
		expect(pdf.content).toContain('(Commented [D2R1]: ) Tj');
		// Formatting changes, worded as Word words them.
		expect(pdf.content).toContain('(Formatted:) Tj');
		expect(pdf.content).toMatch(/\( Centered\) Tj/);
		expect(pdf.content).toMatch(/\( Font: (Bold)?\) Tj/);
		// Deletions stay in the text by default.
		expect(pdf.content).toMatch(/\( ?gone ?\) Tj/);
		expect(pdf.content).not.toContain('(Deleted: ) Tj');
		// The dotted connectors and the pane.
		expect(pdf.content).toContain('0.6 w 0 J 0 j [1 1] 0 d');
		expect(pdf.content).toMatch(/0\.94902 0\.94902 0\.94902 rg\n732 0 [\d.]+ 1056 re\nf/);
		// The comment stays a PDF note, placed on its balloon in the pane; links follow the shrink.
		const note = /\/Subtype \/Text \/Rect \[([\d. ]+)\]/.exec(pdf.raw);
		expect(note).not.toBeNull();
		const [left] = note![1]!.split(' ').map(Number);
		expect(left!).toBeGreaterThan(732 * 0.75 * frame.scale);
		expect(pdf.raw).toMatch(/\/Subtype \/Text [^\n]*\/T \(Cy Young\)/);
		const link = /\/Subtype \/Link \/Rect \[([\d. ]+)\]/.exec(pdf.raw);
		const [, bottom, right] = link![1]!.split(' ').map(Number);
		expect(right!).toBeLessThan(720 * 0.75 * frame.scale + 1);
		expect(bottom!).toBeGreaterThan(frame.offsetYPt);
		// Formatting changes get change bars too.
		expect(pdf.content).toMatch(/\n72 [\d.]+ 1 [\d.]+ re\n72 [\d.]+ 1 [\d.]+ re\n/);
	});

	it('moves deletions to balloons with every revision in balloons', async () => {
		const { bytes } = await exportDocxPdf(await reviewed(), {
			measurer,
			markup: 'all',
			markupBalloons: 'all',
		});
		await dump('balloons-all', bytes);
		const pdf = inspect(bytes);
		expect(pdf.content).toContain('(Deleted: ) Tj');
		// Written once: in its balloon, not struck through in the text.
		expect(pdf.content.match(/gone/g)).toHaveLength(1);
		expect(pdf.content).toContain('(Formatted:) Tj');
		expect(pdf.content).toMatch(/\(added ?\) Tj/);
	});

	it('shows only comments with revisions inline, and no pane without balloons', async () => {
		const model = await reviewed();
		const comments = inspect(
			(await exportDocxPdf(model, { measurer, markup: 'all', markupBalloons: 'comments' })).bytes,
		);
		expect(comments.content).toContain('(Commented [CY1]: ) Tj');
		expect(comments.content).not.toContain('Formatted');
		const none = inspect(
			(await exportDocxPdf(model, { measurer, markup: 'all', markupBalloons: 'none' })).bytes,
		);
		expect(none.content).not.toContain('Commented');
		expect(none.content).not.toMatch(/^[\d.]+ 0 0 [\d.]+ 0 [\d.]+ cm\n[\d.]+ 0 0 -/);
		expect(none.raw).toContain('/Subtype /Text');
		const clean = inspect((await exportDocxPdf(model, { measurer })).bytes);
		expect(clean.content).not.toContain('Commented');
		expect(clean.content).not.toContain('Formatted');
	});
});

/** Writes a PDF for manual checks when PDF_DUMP names a folder. */
async function dump(name: string, bytes: Uint8Array) {
	const folder = process.env.PDF_DUMP;
	if (folder) await writeFile(`${folder}/${name}.pdf`, bytes);
}
