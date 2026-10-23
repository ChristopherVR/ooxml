import { readFileSync } from 'node:fs';
import path from 'node:path';
import { describe, expect, it, vi } from 'vitest';
import { NS, parseXml } from '../../xml/index';
import { RELATIONSHIP_TYPES } from '../../opc/index';
import { readZipParts, SourceIndex } from './package';
import { parseRichString, parseSharedStrings } from './shared-strings';
import { rootChildren } from './sheet-data-source';
import { xChildren } from './xml-util';

// The DOM parser is watched, so a test can tell which path read a part.
vi.mock(import('../../xml/index'), async (importOriginal) => {
	const actual = await importOriginal();
	return { ...actual, parseXml: vi.fn(actual.parseXml) };
});

const PALETTE: string[] = [];

/** The strings as the DOM path reads them. */
const domStrings = (xml: string) =>
	xChildren(parseXml(xml).documentElement, 'si').map((si) => parseRichString(si, PALETTE));

/** Reads `xml` and says whether the DOM parser was used. */
function read(xml: string, maxXmlElements?: number) {
	vi.mocked(parseXml).mockClear();
	const strings = parseSharedStrings(xml, PALETTE, maxXmlElements);
	return { strings, dom: vi.mocked(parseXml).mock.calls.length > 0 };
}

const sst = (items: string, root = `xmlns="${NS.x}"`, prefix = '') =>
	`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<${prefix}sst ${root} count="9" uniqueCount="9">${items}</${prefix}sst>`;

const ITEMS = [
	'<si><t>plain</t></si>',
	'<si><t xml:space="preserve">  spaced &amp; &lt;escaped&gt; &#x41;&#66;\r\nline </t></si>',
	'<si><r><rPr><b/><sz val="12"/><color rgb="FFFF0000"/><rFont val="Arial"/></rPr><t>Bold</t></r><r><t xml:space="preserve"> tail_x000D_</t></r></si>',
	'<si><t>漢字</t><rPh sb="0" eb="2"><t>かんじ</t></rPh><phoneticPr fontId="1"/></si>',
	'<si><t><![CDATA[a < b & "c"]]></t></si>',
	'<!-- between strings --><si><!-- inside --><t>com<!-- in text -->ment</t></si>',
	'<si><t/></si>',
	'<si/>',
	'<si><t>last</t></si>',
].join('');

describe('parseSharedStrings', () => {
	it('reads the strings with the light reader, as the DOM path reads them', () => {
		for (const xml of [
			sst(ITEMS),
			sst(
				ITEMS.replace(/<(\/?)(si|t|r|rPr|rPh|b|sz|color|rFont|phoneticPr)\b/g, '<$1x:$2'),
				`xmlns:x="${NS.x}"`,
				'x:',
			),
			sst(ITEMS, `xmlns="${NS.x}"`).replace('?>\r\n', '?>\r\n<!-- written by a tool -->'),
		]) {
			expect(rootChildren(xml)).toBeDefined();
			const { strings, dom } = read(xml);
			expect(dom).toBe(false);
			expect(strings).toEqual(domStrings(xml));
			expect(strings).toHaveLength(9);
		}
		const { strings } = read(sst(ITEMS));
		expect(strings[1]?.text).toBe('  spaced & <escaped> AB\nline ');
		expect(strings[2]?.runs?.[0]?.font?.bold).toBe(true);
		expect(strings[2]?.text).toBe('Bold tail\r');
		expect(strings[3]?.text).toBe('漢字');
		expect(strings[4]?.text).toBe('a < b & "c"');
		expect(strings[5]?.text).toBe('comment');
		expect(strings.slice(6, 8).map((s) => s.text)).toEqual(['', '']);
	});

	it.each([
		'excel-features.xlsx',
		'excel-smartart.xlsx',
		'excel-chart-families.xlsx',
		'excel-pivot-subtotals-all.xlsx',
	])('reads the shared strings of %s as the DOM path does', async (name) => {
		const bytes = readFileSync(path.join(import.meta.dirname, '..', '__fixtures__', name));
		const source = new SourceIndex(await readZipParts(new Uint8Array(bytes)));
		const part = source.targetOfType(source.workbookPart() ?? '', RELATIONSHIP_TYPES.sharedStrings);
		const xml = (part && source.text(part)) || '';
		expect(xml).toContain('<si>');
		const { strings, dom } = read(xml);
		expect(dom).toBe(false);
		expect(strings).toEqual(domStrings(xml));
	});

	it.each([
		[
			'a namespace declaration inside',
			(xml: string) => xml.replace('<si><t>last', '<si xmlns:y="urn:y"><t>last'),
		],
		[
			'a processing instruction after the last string',
			(xml: string) => xml.replace('</sst>', '<?pi?></sst>'),
		],
		['a processing instruction after the root', (xml: string) => `${xml}<?pi?>`],
		['a self-closing root', () => `<sst xmlns="${NS.x}" count="0"/>`],
	])('leaves a part with %s to the DOM parser, with the same strings', (_, edit) => {
		const xml = edit(sst(ITEMS));
		const { strings, dom } = read(xml);
		expect(dom).toBe(true);
		expect(strings).toEqual(domStrings(xml));
	});

	it('still refuses declarations, also in a comment outside the root', () => {
		for (const xml of [
			sst(ITEMS).replace('?>\r\n', '?>\r\n<!-- <!DOCTYPE x> -->'),
			sst(ITEMS).replace('<t>last</t>', '<t><![CDATA[<!ENTITY x "y">]]></t>'),
		])
			expect(() => parseSharedStrings(xml, PALETTE)).toThrow(/DTD or entity declarations/);
	});

	it('applies the element and nesting limits the DOM parser applies', () => {
		const xml = sst(ITEMS);
		// The root, nine `<si>` and the elements inside them.
		const elements = (xml.replace(/<!\[CDATA\[[^]*?\]\]>/g, '').match(/<(?![!?/])/g) ?? []).length;
		expect(read(xml, elements).dom).toBe(false);
		expect(() => parseSharedStrings(xml, PALETTE, elements - 1)).toThrow(
			expect.objectContaining({
				name: 'WorkbookTooLargeError',
				limit: 'maxXmlElements',
				maximum: elements - 1,
				message: `XLSX shared strings XML exceeds the limit of ${elements - 1} elements`,
			}),
		);
		const deep = sst(`<si><t>${'<r>'.repeat(600)}x${'</r>'.repeat(600)}</t></si>`);
		expect(() => parseSharedStrings(deep, PALETTE)).toThrow(
			expect.objectContaining({ name: 'XmlLimitError', limit: 'maxDepth' }),
		);
	});
});
