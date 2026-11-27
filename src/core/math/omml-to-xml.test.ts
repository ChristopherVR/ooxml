import { describe, expect, it } from 'vitest';
import { latexToOmmlXml } from './omml-to-xml';

describe('latexToOmmlXml schema order', () => {
	it('writes run properties before the text and a hidden degree for a square root', () => {
		const xml = latexToOmmlXml(String.raw`\sqrt{x} + \text{ab} + \sin x + \sqrt[3]{y}`, false);
		// CT_R: m:rPr precedes m:t.
		expect(xml).toContain('<m:r><m:rPr><m:nor m:val="1"/></m:rPr><m:t>ab</m:t></m:r>');
		expect(xml).toContain('<m:r><m:rPr><m:sty m:val="p"/></m:rPr><m:t>sin</m:t></m:r>');
		expect(xml).not.toMatch(/<\/m:t><m:rPr>/);
		// CT_Rad: m:radPr, then m:deg (empty when hidden, as Word writes it), then m:e.
		expect(xml).toContain(
			'<m:rad><m:radPr><m:degHide m:val="1"/></m:radPr><m:deg/><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad>',
		);
		expect(xml).toContain('<m:rad><m:deg><m:r><m:t>3</m:t></m:r></m:deg><m:e>');
		expect(xml).not.toContain('</m:radPr><m:e>');
	});
});
