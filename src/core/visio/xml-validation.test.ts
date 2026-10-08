import { describe, expect, it } from 'vitest';
import { DEFAULTS } from './package-common';
import { inspectXml } from './xml-validation';

const inspect = (xml: string) => inspectXml(xml, DEFAULTS, () => {});

describe('inspectXml namespace scoping', () => {
	it('resolves prefixes declared on the element or an ancestor', () => {
		expect(() => inspect('<r xmlns:a="urn:a"><c a:x="1"/></r>')).not.toThrow();
		expect(() => inspect('<r><c xmlns:a="urn:a" a:x="1"/></r>')).not.toThrow();
		expect(() => inspect('<r xml:space="preserve"/>')).not.toThrow();
	});

	it('rejects a prefix bound only on a sibling or not at all', () => {
		expect(() => inspect('<r><c xmlns:a="urn:a"/><d a:x="1"/></r>')).toThrow(/Unbound/);
		expect(() => inspect('<r __proto__:x="1"/>')).toThrow(/Unbound/);
		expect(() => inspect('<r toString:x="1"/>')).toThrow(/Unbound/);
	});

	it('treats a __proto__ prefix as an ordinary binding', () => {
		expect(() => inspect('<r xmlns:__proto__="urn:p" __proto__:x="1"/>')).not.toThrow();
	});

	it('rejects attributes that expand to the same name through different prefixes', () => {
		expect(() => inspect('<r xmlns:a="urn:x" xmlns:b="urn:x" a:n="1" b:n="2"/>')).toThrow(
			/Duplicate expanded/,
		);
	});
});
