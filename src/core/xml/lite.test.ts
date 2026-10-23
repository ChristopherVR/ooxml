import { describe, expect, it } from 'vitest';
import { LiteDeclinedError, LiteElement, liteChildElements, parseLiteFragment } from './lite';
import { elements, parseXml, type XmlElement } from './xml';

interface Shape {
	name: string;
	ns: string | null;
	attrs: Record<string, string>;
	text: string;
	children: Shape[];
}

const X = 'urn:main';

/** The parts of an element the readers look at, from either implementation. */
function shape(element: XmlElement): Shape {
	const attrs: Record<string, string> = {};
	if (element instanceof LiteElement) {
		for (let i = 0; i < element.attrs.length; i += 2)
			attrs[element.attrs[i] as string] = element.attrs[i + 1] as string;
	} else {
		for (const a of Array.from(element.attributes))
			if (!a.name.startsWith('xmlns')) attrs[a.name] = a.value;
	}
	return {
		name: element.localName,
		ns: element.namespaceURI,
		attrs,
		text: element.textContent ?? '',
		children: elements(element).map(shape),
	};
}

/** Fragments the light reader leaves to the full parser. */
const DECLINED: [string, string][] = [
	['an undeclared element prefix', '<y:row/>'],
	['a namespace declaration', '<row xmlns="urn:other"/>'],
	['a prefixed namespace declaration', '<row xmlns:y="urn:y"/>'],
	['a comment with -- inside', '<row/><!-- a -- b -->'],
	['a comment ending in -', '<row/><!-- a --->'],
	['an unterminated comment', '<row/><!-- note'],
	['unterminated CDATA', '<t><![CDATA[x</t>'],
	['CDATA quoting a DTD', '<t><![CDATA[<!DOCTYPE x>]]></t>'],
	['a comment quoting an entity declaration', '<row/><!-- <!ENTITY x "y"> -->'],
	['a declaration', '<!DOCTYPE x><row/>'],
	['a processing instruction', '<?pi x?><row/>'],
	['an unknown entity', '<t>&nbsp;</t>'],
	['a bare ampersand', '<t>a & b</t>'],
	['an unclosed element', '<row><c>'],
	['a mismatched end tag', '<row></c>'],
	['a stray end tag', '</row>'],
	['an unquoted attribute', '<row r=1/>'],
	['a duplicate attribute', '<row r="1" r="2"/>'],
	['a < in an attribute', '<row r="<"/>'],
	['a missing space between attributes', '<row r="1"s="2"/>'],
	['an undeclared attribute prefix', '<row y:a="1"/>'],
];

const lite = (inner: string, namespaces: Record<string, string> = { '': X }) =>
	parseLiteFragment(inner, new LiteElement('data', 'data', X), namespaces);

/** Parses `inner` both ways (inside a root declaring `decl`) and returns the two shapes. */
function both(inner: string, decl = `xmlns="${X}"`) {
	const dom = parseXml(`<root ${decl}><data>${inner}</data></root>`).documentElement;
	const data = elements(dom)[0] as XmlElement;
	const namespaces: Record<string, string> = {};
	for (const m of decl.matchAll(/xmlns(?::(\w+))?="([^"]*)"/g)) namespaces[m[1] ?? ''] = m[2] ?? '';
	const parsed = lite(inner, namespaces);
	return {
		dom: shape(data).children,
		lite: parsed ? shape(parsed.asElement()).children : undefined,
	};
}

describe('parseLiteFragment', () => {
	it('reads elements, attributes and text like the DOM parser', () => {
		const { dom, lite: light } = both(
			`<row r="1" spans="1:3"><c r="A1" t="inlineStr"><is><t xml:space="preserve"> a &amp; b </t></is></c><c r="B1"><f>SUM(A1:A2)&gt;0</f><v>3</v></c><c r="C1"/></row>`,
		);
		expect(light).toEqual(dom);
		expect(dom[0]?.children[0]?.children[0]?.text).toBe(' a & b ');
	});

	it('decodes character references and normalizes line endings and attribute whitespace', () => {
		const inner = `<t a="x\ty\nz&#10;">l1\r\nl2\rl3&#13;&#x41;&#66;&apos;&quot;&lt;</t>`;
		const { dom, lite: light } = both(inner);
		expect(light).toEqual(dom);
		expect(light?.[0]?.text).toBe('l1\nl2\nl3\rAB\'"<');
		expect(light?.[0]?.attrs['a']).toBe('x y z\n');
	});

	it('resolves element prefixes through the namespaces in scope', () => {
		const { dom, lite: light } = both(
			`<x:row r="2"><x:c r="A2"><x:v>1</x:v></x:c><c/></x:row>`,
			`xmlns:x="${X}"`,
		);
		expect(light).toEqual(dom);
		expect(light?.[0]?.ns).toBe(X);
		expect(light?.[0]?.children[1]?.ns).toBeNull();
	});

	it('keeps prefixed attributes under their qualified name', () => {
		const parsed = lite('<row r="1" x14ac:dyDescent="0.25" xml:space="preserve"/>', {
			'': X,
			x14ac: 'urn:x14ac',
		});
		const row = parsed?.childNodes[0] as LiteElement;
		expect(row.getAttribute('x14ac:dyDescent')).toBe('0.25');
		expect(row.hasAttribute('r')).toBe(true);
		expect(row.hasAttribute('s')).toBe(false);
		expect(row.getAttribute('s')).toBeNull();
	});

	it('skips comments and reads CDATA text like the DOM parser', () => {
		const inner = [
			'<!-- before --><row r="1"><!-- in a row -->',
			'<c r="A1" t="inlineStr"><!-- in a cell --><is><t>a<!-- in text -->b</t></is></c>',
			'<c r="B1"><f><![CDATA[IF(A1<2,"x & y","]]")]]></f><v>1<![CDATA[]]>2</v></c>',
			'<c r="C1" t="inlineStr"><is><t><![CDATA[ line\r\nnext <b> ]]>tail &amp; more</t></is></c>',
			'</row><!-- <row r="9"/> --><row r="2"/><!---->',
		].join('');
		const { dom, lite: light } = both(inner);
		expect(light).toEqual(dom);
		expect(dom[0]?.children[2]?.children[0]?.children[0]?.text).toBe(' line\nnext <b> tail & more');
		expect(dom.map((row) => row.attrs.r)).toEqual(['1', '2']);
	});

	it.each(DECLINED)('declines %s', (_, inner) => {
		expect(lite(inner)).toBeUndefined();
	});
});

describe('liteChildElements', () => {
	it('yields the elements parseLiteFragment reads, one at a time', () => {
		const inner = [
			'<!-- before --><row r="1"><c r="A1"><v>1</v></c></row>\n',
			'<row r="2"/><![CDATA[text]]><x:row r="3"><c><![CDATA[a<b]]></c></x:row>',
		].join('');
		const namespaces = { '': X, x: X };
		const whole = parseLiteFragment(inner, new LiteElement('data', 'data', X), namespaces);
		const rows = liteChildElements(inner, namespaces);
		const first = rows.next();
		expect(first.done).toBe(false);
		expect(shape(first.value as XmlElement)).toEqual(shape(elements(whole!.asElement())[0]!));
		expect([...rows].map(shape)).toEqual(elements(whole!.asElement()).slice(1).map(shape));
	});

	it.each(DECLINED)('throws a LiteDeclinedError for %s', (_, inner) => {
		expect(() => [...liteChildElements(inner, { '': X })]).toThrow(LiteDeclinedError);
	});
});
