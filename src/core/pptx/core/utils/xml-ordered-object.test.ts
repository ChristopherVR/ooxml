import { describe, expect, it } from 'vitest';

import { PptxRuntimeDependencyFactory } from '../core/factories/PptxRuntimeDependencyFactory';
import { orderedXmlObject } from './xml-ordered-object';

const builder = new PptxRuntimeDependencyFactory().createBuilder();
const roundTrip = (xml: string) => builder.build(orderedXmlObject(xml)) as string;

describe('orderedXmlObject', () => {
	it('keeps interleaved siblings in document order through the runtime builder', () => {
		const xml =
			'<c:plotArea><c:barChart><c:ser>a</c:ser></c:barChart><c:lineChart><c:ser>b</c:ser></c:lineChart><c:barChart><c:ser>c</c:ser></c:barChart><c:catAx/><c:valAx/><c:catAx/></c:plotArea>';
		expect(roundTrip(xml)).toBe(
			'<c:plotArea><c:barChart><c:ser>a</c:ser></c:barChart><c:lineChart><c:ser>b</c:ser></c:lineChart><c:barChart><c:ser>c</c:ser></c:barChart><c:catAx></c:catAx><c:valAx></c:valAx><c:catAx></c:catAx></c:plotArea>',
		);
	});

	it('keeps adjacent repeats as one array and runs around a break in order', () => {
		const tree = orderedXmlObject('<a:p><a:r><a:t>A</a:t></a:r><a:r><a:t>B</a:t></a:r></a:p>');
		expect(tree).toStrictEqual({ 'a:p': { 'a:r': [{ 'a:t': 'A' }, { 'a:t': 'B' }] } });
		expect(roundTrip('<a:p><a:r><a:t>A</a:t></a:r><a:br/><a:r><a:t>B</a:t></a:r></a:p>')).toBe(
			'<a:p><a:r><a:t>A</a:t></a:r><a:br></a:br><a:r><a:t>B</a:t></a:r></a:p>',
		);
	});

	it('decodes entities once and keeps text whitespace and attributes', () => {
		const xml = '<c:v xml:space="preserve"> a &amp; &lt;b&gt; &#x41; </c:v>';
		expect(orderedXmlObject(xml)).toStrictEqual({
			'c:v': { '@_xml:space': 'preserve', '#text': ' a & <b> A ' },
		});
		// The runtime builder encodes text as every part it writes (`encodeXmlTextValue`).
		expect(roundTrip('<c:tx><c:v>a &amp; "b"</c:v></c:tx>')).toBe(
			'<c:tx><c:v>a &amp; &quot;b&quot;</c:v></c:tx>',
		);
	});

	it('refuses a tag name that would reach the prototype', () => {
		expect(() => orderedXmlObject('<r><__proto__>x</__proto__></r>')).toThrow();
	});
});
