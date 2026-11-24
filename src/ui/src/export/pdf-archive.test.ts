import { describe, expect, it } from 'vitest';
import { syntheticTrueType } from '../pptx/export/__fixtures__/synthetic-truetype';
import { infoText, pdfALevel, xmpMetadata } from './pdf-archive';
import { PdfEmbeddedFont, PdfFontSet } from './pdf-embedded-font';
import { buildVectorPdf, PdfVectorPage } from './pdf-vector';

const latin1 = (bytes: Uint8Array) => Buffer.from(bytes).toString('latin1');
const utf8 = (bytes: Uint8Array) => Buffer.from(bytes).toString('utf8');

/** The info dictionary's `/Title`, decoded from a literal or UTF-16BE hex string. */
function infoTitle(raw: string): string {
	const hex = /\/Title <FEFF([0-9A-F]*)>/.exec(raw);
	if (hex)
		return String.fromCharCode(...hex[1]!.match(/.{4}/g)!.map((code) => Number.parseInt(code, 16)));
	return /\/Title \(((?:\\.|[^\\)])*)\)/.exec(raw)![1]!.replace(/\\(.)/g, '$1');
}

const xmpTitle = (raw: string) =>
	/<dc:title><rdf:Alt><rdf:li xml:lang="x-default">([^<]*)</
		.exec(raw)![1]!
		.replace(/&lt;/g, '<')
		.replace(/&gt;/g, '>')
		.replace(/&quot;/g, '"')
		.replace(/&amp;/g, '&');

describe('PDF/A archive parts', () => {
	it('keeps the info dictionary and XMP equal for characters XML cannot hold', () => {
		// ISO 19005-2 6.6.2.3: a control character or an unpaired surrogate made them differ.
		expect(infoText('A\u0001B\u000bC\ufffeD\tE')).toBe('ABCD\tE');
		expect(infoText('x\ud800y\udc00z\ud83d\ude00')).toBe('x\ufffdy\ufffdz\ud83d\ude00');
		const page = new PdfVectorPage(100, 100);
		const title = 'R\u00e9sum\u00e9 <a> & "b"\u0001 \ud800';
		const bytes = buildVectorPdf([page], { title, pdfA: true, date: new Date(0) });
		const raw = latin1(bytes);
		expect(infoTitle(raw)).toBe('R\u00e9sum\u00e9 <a> & "b" \ufffd');
		expect(xmpTitle(utf8(bytes))).toBe(infoTitle(raw));
	});

	it('declares level U when every font is embedded and B when a standard font is used', () => {
		const font = PdfEmbeddedFont.create({ family: 'Synth', data: syntheticTrueType() })!;
		const embedded = new PdfVectorPage(100, 100);
		embedded.embeddedText(font, 12, 10, 20, 'AB');
		expect(pdfALevel([embedded])).toBe('U');
		const xmp = latin1(buildVectorPdf([embedded], { pdfA: true, date: new Date(0) }));
		expect(xmp).toContain('<pdfaid:part>2</pdfaid:part><pdfaid:conformance>U</pdfaid:conformance>');
		const standard = new PdfVectorPage(100, 100);
		standard.text('Helvetica', 12, 10, 20, 'AB');
		expect(pdfALevel([embedded, standard])).toBe('B');
		expect(latin1(buildVectorPdf([standard], { pdfA: true }))).toContain(
			'<pdfaid:conformance>B</pdfaid:conformance>',
		);
		expect(utf8(xmpMetadata({}, new Date(0)))).toContain(
			'<pdfaid:conformance>B</pdfaid:conformance>',
		);
	});

	it('writes text in any family with a supplied font when every font must be embedded', () => {
		const files = [{ family: 'Synth', data: syntheticTrueType() }];
		expect(new PdfFontSet(files).prefers('Segoe UI')).toBe(false);
		expect(new PdfFontSet(files).prefers('Synth')).toBe(true);
		expect(new PdfFontSet(files, { embedAll: true }).prefers('Segoe UI')).toBe(true);
		// Nothing to embed with: the standard fonts stay the fallback.
		expect(new PdfFontSet([], { embedAll: true }).prefers('Segoe UI')).toBe(false);
	});
});
