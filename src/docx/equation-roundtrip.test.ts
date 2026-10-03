import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { loadDocx } from './parse.js';
import type { Paragraph } from './model.js';
import { parseXml } from './xml.js';
import { parseOmml, convertOmmlToMathMl } from '../math/index.js';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
const inline =
	'<m:oMath ext:keep="yes"><m:f><m:num><m:r><m:t>x</m:t></m:r></m:num><m:den><m:r><m:t>2</m:t></m:r></m:den></m:f></m:oMath>';
const display =
	'<m:oMathPara><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr><m:oMath><m:r><m:t>y</m:t></m:r></m:oMath></m:oMathPara>';
async function fixture(math = inline) {
	const zip = new JSZip();
	zip.file(
		'word/document.xml',
		`<w:document xmlns:w="${W}" xmlns:m="${M}" xmlns:ext="urn:test"><w:body><w:p><w:r><w:t>Before</w:t></w:r>${math}<w:r><w:t>After</w:t></w:r></w:p><w:sectPr/></w:body></w:document>`,
	);
	return zip.generateAsync({ type: 'uint8array' });
}
const paragraph = (loaded: Awaited<ReturnType<typeof loadDocx>>) =>
	loaded.model.blocks[0] as Paragraph;
async function xmlOf(bytes: Uint8Array) {
	return (await JSZip.loadAsync(bytes)).file('word/document.xml')!.async('string');
}

describe('imported Word equations', () => {
	it('exposes independently parseable inline and display OMML in document order', async () => {
		for (const [source, isDisplay] of [
			[inline, false],
			[display, true],
		] as const) {
			const bytes = await fixture(source);
			const loaded = await loadDocx(bytes);
			const runs = paragraph(loaded).runs;
			expect(runs.map((run) => run.text)).toEqual(['Before', '', 'After']);
			expect(runs[1]!.equation?.display).toBe(isDisplay);
			expect(parseXml(runs[1]!.equation!.omml).documentElement.namespaceURI).toBe(M);
			expect(convertOmmlToMathMl(parseOmml(runs[1]!.equation!.omml))).toContain(
				isDisplay ? 'y' : 'mfrac',
			);
			expect(await loaded.save()).toEqual(bytes);
		}
	});

	it('preserves original equation subtree while surrounding runs split and change', async () => {
		const loaded = await loadDocx(await fixture());
		const p = paragraph(loaded);
		const equation = p.runs[1]!;
		p.runs = [{ text: 'New ' }, { text: 'before', bold: true }, equation, { text: 'after' }];
		const bytes = await loaded.save();
		const xml = await xmlOf(bytes);
		expect(xml).toContain(inline);
		expect(xml.indexOf('before')).toBeLessThan(xml.indexOf('<m:oMath'));
		expect(xml.indexOf('<m:oMath')).toBeLessThan(xml.indexOf('after'));
		expect(paragraph(await loadDocx(bytes)).runs[2]!.equation).toEqual(equation.equation);
	});

	it('keeps equations inside tracked-change wrappers when editing neighboring text', async () => {
		const loaded = await loadDocx(
			await fixture(`<w:ins w:id="9" w:author="Editor">${inline}</w:ins>`),
		);
		expect(paragraph(loaded).runs[1]!.revision?.kind).toBe('insert');
		paragraph(loaded).runs[0]!.text = 'Changed';
		const xml = await xmlOf(await loaded.save());
		expect(xml).toContain(inline);
		expect(xml).toContain('w:author="Editor"');
	});

	it('rejects source changes and new equations instead of silently replacing them', async () => {
		const loaded = await loadDocx(await fixture());
		paragraph(loaded).runs[1]!.equation!.omml = display;
		await expect(loaded.save()).rejects.toThrow('Equation editing or insertion is unsupported');
		const plain = await loadDocx(await fixture(''));
		paragraph(plain).runs.push({ text: '', equation: { omml: inline, display: false } });
		await expect(plain.save()).rejects.toThrow('Equation editing or insertion is unsupported');
	});

	it('preserves equations in simple fields and hyperlinks instead of treating them as empty text', async () => {
		for (const wrapper of [
			`<w:fldSimple w:instr=" EQ ">${inline}</w:fldSimple>`,
			`<w:hyperlink w:anchor="target">${inline}</w:hyperlink>`,
		]) {
			const loaded = await loadDocx(await fixture(wrapper));
			expect(paragraph(loaded).runs[1]!.equation).toBeDefined();
			paragraph(loaded).runs[0]!.text = 'Changed';
			expect(await xmlOf(await loaded.save())).toContain(inline);
		}
	});

	it('keeps repeated identical equations and distinct source metadata across run splits', async () => {
		const distinct = inline.replace('ext:keep="yes"', 'ext:keep="second"');
		const loaded = await loadDocx(await fixture(inline + inline + distinct));
		const p = paragraph(loaded);
		p.runs.unshift({ text: 'Added' });
		const xml = await xmlOf(await loaded.save());
		expect(xml.match(/<m:oMath ext:keep="yes">/g)).toHaveLength(2);
		expect(xml.match(/<m:oMath ext:keep="second">/g)).toHaveLength(1);
		expect(
			paragraph(await loadDocx(await loaded.save())).runs.filter((run) => run.equation),
		).toHaveLength(3);
	});
});
