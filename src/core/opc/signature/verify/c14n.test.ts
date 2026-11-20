import { DOMParser } from '@xmldom/xmldom';
import { SignedXml } from 'xml-crypto';
import { describe, expect, it } from 'vitest';

import {
	EXCLUSIVE_C14N,
	EXCLUSIVE_C14N_WITH_COMMENTS,
	INCLUSIVE_C14N,
	INCLUSIVE_C14N_WITH_COMMENTS,
} from '../constants';
import { canonicalize } from './c14n';

const parse = (xml: string): Document =>
	new DOMParser().parseFromString(xml, 'text/xml') as unknown as Document;

const ALGORITHMS = [
	INCLUSIVE_C14N,
	INCLUSIVE_C14N_WITH_COMMENTS,
	EXCLUSIVE_C14N,
	EXCLUSIVE_C14N_WITH_COMMENTS,
];

/** xml-crypto, a widely used implementation, as the reference for element subsets. */
function reference(element: Element, algorithm: string): string {
	const ancestorNamespaces: { prefix: string; namespaceURI: string }[] = [];
	const seen = new Set<string>([element.prefix ?? '']);
	for (const attribute of Array.from(element.attributes)) {
		if (attribute.name === 'xmlns') seen.add('');
		else if (attribute.name.startsWith('xmlns:')) seen.add(attribute.name.slice(6));
	}
	for (let node = element.parentNode; node?.nodeType === 1; node = node.parentNode) {
		for (const attribute of Array.from((node as Element).attributes)) {
			const prefix =
				attribute.name === 'xmlns'
					? ''
					: attribute.name.startsWith('xmlns:')
						? attribute.name.slice(6)
						: undefined;
			if (prefix === undefined || seen.has(prefix)) continue;
			seen.add(prefix);
			if (attribute.value) ancestorNamespaces.push({ prefix, namespaceURI: attribute.value });
		}
	}
	// eslint-disable-next-line @typescript-eslint/no-explicit-any
	return new SignedXml().getCanonXml([algorithm], element as any, { ancestorNamespaces });
}

const SAMPLES = [
	'<a xmlns="urn:a" xmlns:b="urn:b" z="1" b:y="2" a="3"><b:c>t &amp; &lt; &gt;</b:c><!--x--><d/></a>',
	'<r xmlns:p="urn:p" xmlns:q="urn:q"><p:x q:attr="v"><y xmlns="urn:d"><z xmlns=""/></y></p:x></r>',
	'<Signature xmlns="http://www.w3.org/2000/09/xmldsig#"><Object Id="o"><m:M xmlns:m="urn:m" Id="i">\n  <m:V a="&#9;&#10;&#13;&quot;">x</m:V></m:M></Object></Signature>',
	'<a><?pi data?><b><![CDATA[<c>&]]></b>\r\n<e attr="a&lt;b"/></a>',
	'<a xmlns:x="urn:x" xmlns:unused="urn:u"><x:b xmlns:x="urn:x"><x:c xmlns:x="urn:y"/></x:b></a>',
];

describe('canonicalize', () => {
	it.each(SAMPLES.flatMap((xml) => ALGORITHMS.map((algorithm) => [xml, algorithm] as const)))(
		'renders every element of %s like the reference under %s',
		(xml, algorithm) => {
			const document = parse(xml);
			for (const element of Array.from(document.getElementsByTagName('*')))
				expect(canonicalize(element, algorithm)).toBe(reference(element, algorithm));
		},
	);

	it('renders a document with its processing instructions and, on request, comments', () => {
		const document = parse('<?xml version="1.0"?><?pi a?><!--c--><r b="1" a="2"/><!--d-->');
		expect(canonicalize(document, INCLUSIVE_C14N)).toBe('<?pi a?>\n<r a="2" b="1"></r>');
		expect(canonicalize(document, INCLUSIVE_C14N_WITH_COMMENTS)).toBe(
			'<?pi a?>\n<!--c-->\n<r a="2" b="1"></r>\n<!--d-->',
		);
	});

	it('gives the apex the xml: attributes it inherits under Canonical XML 1.0 only', () => {
		const apex = parse('<r xml:lang="en" xml:space="preserve"><a xml:lang="de"/></r>')
			.documentElement.firstChild as Element;
		expect(canonicalize(apex, INCLUSIVE_C14N)).toBe('<a xml:lang="de" xml:space="preserve"></a>');
		expect(canonicalize(apex, EXCLUSIVE_C14N)).toBe('<a xml:lang="de"></a>');
	});

	it('refuses an algorithm it does not implement', () => {
		expect(() => canonicalize(parse('<a/>'), 'urn:other')).toThrow(/Unsupported/);
	});
});
