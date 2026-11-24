import { readFile } from 'node:fs/promises';
import { describe, expect, it } from 'vitest';
import { decodeFilters } from './filters';
import { openPdf, PdfPasswordError } from './index';
import { PdfLexer } from './lexer';
import { PdfName, PdfRef, PdfString } from './objects';
import { buildPdf, simplePdf, stream } from './test-support/build';

const fixture = (name: string) =>
	readFile(new URL(`../docx/pdf-import/__fixtures__/${name}.pdf`, import.meta.url)).then(
		(bytes) => new Uint8Array(bytes),
	);
const text = (spans: { text: string }[]) => spans.map((span) => span.text).join('');

describe('PdfLexer', () => {
	it('reads strings, names, references and nested containers', () => {
		const lexer = new PdfLexer(
			new TextEncoder().encode('<< /A#20B (x\\(y\\)\\101\\\nz) /C [1 0 R <414> -.5] /D true >>'),
		);
		const dict = lexer.object() as Map<string, unknown>;
		expect([...dict.keys()]).toEqual(['A B', 'C', 'D']);
		expect(new TextDecoder().decode((dict.get('A B') as PdfString).bytes)).toBe('x(y)Az');
		const array = dict.get('C') as unknown[];
		expect(array[0]).toEqual(new PdfRef(1, 0));
		expect([...(array[1] as PdfString).bytes]).toEqual([0x41, 0x40]);
		expect(array[2]).toBe(-0.5);
		expect(dict.get('D')).toBe(true);
	});

	it('survives unbalanced and truncated input', () => {
		expect(new PdfLexer(new TextEncoder().encode('[ /A << /B (open')).object()).toEqual([
			new PdfName('A'),
			new Map([['B', new PdfString(new TextEncoder().encode('open'))]]),
		]);
	});
});

describe('stream filters', () => {
	it('undoes ASCII85, ASCIIHex, run length and PNG predictors', () => {
		const dict = (filter: string) =>
			new Map<string, never>([['Filter', new PdfName(filter) as never]]);
		expect(
			new TextDecoder().decode(
				decodeFilters(new TextEncoder().encode('87cURD]i,"Ebo80~>'), dict('ASCII85Decode'), 1e6)
					.bytes,
			),
		).toBe('Hello World!');
		expect(
			new TextDecoder().decode(
				decodeFilters(new TextEncoder().encode('48656C6C6F>'), dict('ASCIIHexDecode'), 1e6).bytes,
			),
		).toBe('Hello');
		expect([
			...decodeFilters(Uint8Array.from([1, 65, 66, 254, 67, 128]), dict('RunLengthDecode'), 1e6)
				.bytes,
		]).toEqual([65, 66, 67, 67, 67]);
		expect(() =>
			decodeFilters(new TextEncoder().encode('z'.repeat(10)), dict('ASCII85Decode'), 8),
		).toThrow(/size limit/);
	});
});

describe('openPdf', () => {
	it('reads text in reading coordinates from a classic file', async () => {
		const reader = await openPdf(simplePdf('BT /F1 12 Tf 72 720 Td (Hello) Tj ( world) Tj ET'));
		const page = await reader.page(0);
		expect(page.width).toBe(612);
		expect(text(page.spans)).toBe('Hello world');
		expect(page.spans[0]).toMatchObject({ x: 72, y: 72, fontSize: 12, angle: 0 });
		expect(page.spans[0]!.font).toMatchObject({ family: 'Arial', bold: false });
		// "Hello" in Helvetica: 722 + 556 + 222 + 222 + 556 thousandths.
		expect(page.spans[0]!.width).toBeCloseTo((2278 * 12) / 1000, 5);
	});

	it('reads object streams and cross-reference streams', async () => {
		const reader = await openPdf(
			simplePdf('BT /F1 10 Tf 1 0 0 1 100 700 Tm (Compressed) Tj ET', { compressed: true }),
		);
		expect(reader.pageCount).toBe(1);
		expect(text((await reader.page(0)).spans)).toBe('Compressed');
	});

	it('rebuilds a damaged cross-reference table by scanning for objects', async () => {
		const bytes = simplePdf('BT /F1 10 Tf 100 700 Td (Damaged) Tj ET');
		const broken = String.fromCharCode(...bytes).replace(/startxref\n\d+/, 'startxref\n999999');
		const reader = await openPdf(Uint8Array.from(broken, (char) => char.charCodeAt(0)));
		expect(text((await reader.page(0)).spans)).toBe('Damaged');
	});

	it('maps codes through ToUnicode, Differences and Identity-H with widths', async () => {
		const toUnicode = stream(
			'',
			'1 begincodespacerange <0000> <FFFF> endcodespacerange 2 beginbfrange <0001> <0002> <0041> <0003> <0003> [<00E9>] endbfrange',
		);
		const reader = await openPdf(
			simplePdf('BT /F2 10 Tf 0 700 Td <000100020003> Tj /F3 10 Tf (\\001\\002) Tj ET', {
				resources: '/Font << /F2 6 0 R /F3 9 0 R >>',
				objects: [
					'<< /Type /Font /Subtype /Type0 /BaseFont /ABCDEF+Calibri-Bold /Encoding /Identity-H /DescendantFonts [7 0 R] /ToUnicode 8 0 R >>',
					'<< /Type /Font /Subtype /CIDFontType2 /BaseFont /Calibri-Bold /DW 500 /W [1 [600 700]] >>',
					toUnicode,
					'<< /Type /Font /Subtype /Type1 /BaseFont /Times-Italic /FirstChar 1 /Widths [500 500] /Encoding << /Differences [1 /fi /eacute] >> >>',
				],
			}),
		);
		const spans = (await reader.page(0)).spans;
		expect(text(spans)).toBe('ABéfié');
		expect(spans[0]!.font).toMatchObject({ family: 'Calibri', bold: true });
		expect(spans[0]!.width).toBeCloseTo((600 + 700 + 500) / 100, 5);
		expect(spans[1]!.font).toMatchObject({ family: 'Times New Roman', italic: true });
	});

	it('places rotated pages, rises and form drawings', async () => {
		const reader = await openPdf(
			buildPdf([
				'<< /Type /Catalog /Pages 2 0 R >>',
				'<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
				'<< /Type /Page /Parent 2 0 R /MediaBox [0 0 600 800] /Rotate 90 /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> /XObject << /X 6 0 R >> >> >>',
				stream('', 'q 1 0 0 1 10 20 cm /X Do Q BT /F1 10 Tf 5 Ts 50 60 Td (up) Tj ET'),
				'<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
				stream(
					'/Type /XObject /Subtype /Form /BBox [0 0 100 100] /Resources << /Font << /F1 5 0 R >> >>',
					'BT /F1 8 Tf 0 0 Td (form) Tj ET',
				),
			]),
		);
		const page = await reader.page(0);
		expect([page.width, page.height]).toEqual([800, 600]);
		const [form, up] = page.spans;
		expect(form).toMatchObject({ text: 'form', x: 20, y: 10 });
		expect(up).toMatchObject({ text: 'up', x: 60, y: 50, rise: 5 });
		expect(up!.angle).toBeCloseTo(-Math.PI / 2, 5);
	});
});

describe('encrypted files (written by pypdf)', () => {
	it('opens an owner-password-only file without a password and reports its permissions', async () => {
		const reader = await openPdf(await fixture('lettre-rc4-owner-only'));
		expect(reader.encrypted).toBe(true);
		expect(reader.permissions).toMatchObject({ print: true, copy: true, modify: false });
		expect(text((await reader.page(0)).spans)).toContain('Corps de la lettre.');
	});

	it.each(['lettre-aes128', 'lettre-aes256'])(
		'opens %s with the user or owner password',
		async (name) => {
			const bytes = await fixture(name);
			await expect(openPdf(bytes)).rejects.toMatchObject({ code: 'password-required' });
			await expect(openPdf(bytes, { password: 'wrong' })).rejects.toBeInstanceOf(PdfPasswordError);
			for (const password of ['secret', 'proprietaire']) {
				const reader = await openPdf(bytes, { password });
				expect(text((await reader.page(0)).spans)).toContain('Cabinet Exemple');
			}
		},
	);
});

describe('a Word PDF', () => {
	it('reads its fonts, pictures, rules and document information', async () => {
		const reader = await openPdf(await fixture('rapport'));
		expect(reader.pageCount).toBe(2);
		expect(reader.tagged).toBe(true);
		expect(reader.info).toMatchObject({ title: 'Rapport annuel', author: 'Exemple' });
		const page = await reader.page(0);
		expect(text(page.spans)).toContain('Un passage en gras');
		expect(page.spans.find((span) => span.text.includes('gras') && span.font.bold)).toBeDefined();
		expect(page.spans.find((span) => span.color === '#c00000')).toBeDefined();
		expect(page.images).toHaveLength(1);
		expect(page.images[0]).toMatchObject({ pixelWidth: 40, pixelHeight: 24 });
		expect(page.images[0]!.format).not.toBe('unsupported');
		// The underline is a thin rule under its word.
		expect(page.rules.some((rule) => rule.y1 - rule.y0 < 1.5 && rule.x1 - rule.x0 > 20)).toBe(true);
	});
});
