import { describe, expect, it } from 'vitest';
import { mergeXmlThreeWay, type XmlMergeOptions } from './merge-three-way';
import { parseXml } from './xml';
const options: XmlMergeOptions = {
	keyElement: (element) => element.getAttribute('id') ?? undefined,
	allowText: () => true,
	allowPositional: () => true,
};
const wrapped = (value: string) => `<root>${value}</root>`;
const content = (result: ReturnType<typeof mergeXmlThreeWay>) =>
	result.status === 'merged' ? parseXml(result.xml).documentElement : undefined;
describe('three-way XML merge', () => {
	it('combines distant edits in a long repeated sequence without changing its order', () => {
		const values = Array.from({ length: 1000 }, (_, i) => `<p>Paragraph ${i}</p>`);
		const local = [...values],
			remote = [...values];
		local[0] = '<p>Local first paragraph</p>';
		remote[999] = '<p>Remote last paragraph</p>';
		const root = content(
			mergeXmlThreeWay(
				wrapped(values.join('')),
				wrapped(local.join('')),
				wrapped(remote.join('')),
				options,
			),
		);
		expect(root?.childNodes.length).toBe(1000);
		expect(root?.firstChild?.textContent).toBe('Local first paragraph');
		expect(root?.lastChild?.textContent).toBe('Remote last paragraph');
	});
	it('combines independent properties and leaf text edits', () => {
		const root = content(
			mergeXmlThreeWay(
				'<root a="1"><t>Hello world</t></root>',
				'<root a="2"><t>Hello brave world</t></root>',
				'<root a="1" b="3"><t>Hello world!</t></root>',
				options,
			),
		);
		expect(root?.getAttribute('a')).toBe('2');
		expect(root?.getAttribute('b')).toBe('3');
		expect(root?.textContent).toBe('Hello brave world!');
	});
	it('retains independent keyed insertions with convergent order', () => {
		const b = wrapped('<x id="0"/><x id="9"/>'),
			l = wrapped('<x id="0"/><x id="2"/><x id="9"/>'),
			r = wrapped('<x id="0"/><x id="1"/><x id="9"/>');
		const left = mergeXmlThreeWay(b, l, r, options),
			right = mergeXmlThreeWay(b, r, l, options);
		expect(left.status).toBe('merged');
		expect(left).toEqual(right);
		expect(content(left)?.getElementsByTagName('x').length).toBe(4);
	});
	it('combines a deletion with an independent keyed edit', () => {
		const merged = content(
			mergeXmlThreeWay(
				wrapped('<x id="a">A</x><x id="b">B</x>'),
				wrapped('<x id="b">B</x>'),
				wrapped('<x id="a">A</x><x id="b">Bee</x>'),
				options,
			),
		);
		expect(merged?.textContent).toBe('Bee');
	});
	it.each([
		['<root a="1"/>', '<root a="2"/>', '<root a="3"/>'],
		[wrapped('<x id="a">A</x>'), wrapped(''), wrapped('<x id="a">B</x>')],
		[wrapped('<x>A</x><x>B</x>'), wrapped('<x>B</x><x>A</x>'), wrapped('<x>A!</x><x>B</x>')],
		[
			wrapped('<x>A</x><x>B</x>'),
			wrapped('<x>A</x><x>B</x><x>C</x>'),
			wrapped('<x>A!</x><x>B</x>'),
		],
	])('declines unsafe overlap or structural matching', (b, l, r) => {
		expect(mergeXmlThreeWay(b, l, r, options).status).toBe('unresolved');
	});
	it('refuses duplicate keyed identities and malicious XML', () => {
		expect(
			mergeXmlThreeWay(
				'<root/>',
				wrapped('<x id="a"/><x id="a"/>'),
				wrapped('<x id="b"/>'),
				options,
			).status,
		).toBe('unresolved');
		expect(
			mergeXmlThreeWay('<root/>', '<!DOCTYPE root><root/>', '<root b="2"/>', options).status,
		).toBe('unresolved');
	});
});
