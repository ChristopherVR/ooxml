import JSZip from 'jszip';
import { afterEach, describe, expect, it, vi } from 'vitest';
import type { Block, Paragraph } from './model';
import { loadDocx } from './parse';
import { docx, p } from './test-support/compare-docs';
import { parseXml, type XmlElement } from './xml';
import { replaceSlots } from './write-slots';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function body(xml: string): { body: XmlElement; byName: (name: string) => XmlElement } {
	const doc = parseXml(`<w:document xmlns:w="${W}"><w:body>${xml}</w:body></w:document>`);
	const element = doc.documentElement!.firstChild as XmlElement;
	const byName = (name: string) =>
		Array.from(element.getElementsByTagNameNS(W, '*')).find(
			(node) => node.getAttribute('n') === name,
		) as XmlElement;
	return { body: element, byName };
}

const order = (parent: XmlElement): string =>
	Array.from(parent.childNodes)
		.map((node) => (node as XmlElement).getAttribute?.('n') ?? (node as XmlElement).localName)
		.join(' ');

afterEach(() => vi.restoreAllMocks());

describe('replaceSlots', () => {
	it('puts each next node in the slot of the old node at its index', () => {
		const { body: parent, byName } = body(
			'<w:p n="a"/><w:bookmarkEnd n="m"/><w:p n="b"/><w:p n="c"/><w:sectPr n="s"/>',
		);
		const [a, b, c, s] = ['a', 'b', 'c', 's'].map(byName) as XmlElement[];
		const x = parent.ownerDocument.createElementNS(W, 'w:p');
		x.setAttribute('n', 'x');
		// Inserting x first shifts every block one slot on; the marker keeps its place.
		replaceSlots(parent, [a!, b!, c!], [x, a!, b!, c!], s!);
		expect(order(parent)).toBe('x m a b c s');
		replaceSlots(parent, [x, a!, b!, c!], [a!, c!], s!);
		expect(order(parent)).toBe('a m c s');
	});

	it('replaces slots inside a content control in place and keeps a repeated node in its last slot', () => {
		const { body: parent, byName } = body(
			'<w:sdt n="sdt"><w:sdtContent n="content"><w:p n="a"/></w:sdtContent></w:sdt><w:p n="b"/><w:p n="c"/>',
		);
		const [a, b, c, content] = ['a', 'b', 'c', 'content'].map(byName) as XmlElement[];
		replaceSlots(parent, [a!, b!, c!], [c!, b!, c!]);
		expect(order(content!)).toBe('');
		expect(order(parent)).toBe('sdt b c');
		replaceSlots(parent, [b!, c!], [c!, a!]);
		expect(order(parent)).toBe('sdt c a');
	});

	it('leaves unchanged slots untouched and moves only what changed', () => {
		const xml = Array.from({ length: 400 }, (_, i) => `<w:p n="p${i}"/>`).join('');
		const { body: parent } = body(`${xml}<w:sectPr/>`);
		const blocks = Array.from(parent.childNodes).slice(0, 400) as XmlElement[];
		const anchor = parent.lastChild as XmlElement;
		const insert = vi.spyOn(parent, 'insertBefore');
		const remove = vi.spyOn(parent, 'removeChild');
		replaceSlots(parent, blocks, blocks, anchor);
		const fresh = parent.ownerDocument.createElementNS(W, 'w:p');
		replaceSlots(parent, blocks, [fresh, ...blocks.slice(1), blocks[0]!], anchor);
		replaceSlots(parent, [fresh, ...blocks.slice(1), blocks[0]!], blocks.slice(1), anchor);
		// One insertion, one move of the first block to the end (a removal and an insertion),
		// and the removal of each.
		expect(insert.mock.calls.length + remove.mock.calls.length).toBeLessThanOrEqual(5);
		expect(Array.from(parent.childNodes).length).toBe(400);
	});
});

describe('saving a long document', () => {
	it('moves no unchanged paragraph when one is inserted at the start', async () => {
		const texts = Array.from({ length: 3000 }, (_, i) => `Paragraph ${i}`);
		const loaded = await loadDocx(await docx(p(...texts)));
		const inserted: Paragraph = { type: 'paragraph', id: 'new', runs: [{ text: 'First' }] };
		const blocks: Block[] = [inserted, ...loaded.model.blocks];
		const element = Object.getPrototypeOf(parseXml(`<a/>`).documentElement!) as Node;
		const moves = [
			vi.spyOn(element, 'insertBefore'),
			vi.spyOn(element, 'removeChild'),
			vi.spyOn(element, 'replaceChild'),
		];
		const saved = await loaded.save({ ...loaded.model, blocks });
		// Changes to the body's children other than appends (reparsing the part appends each one).
		const calls = moves.reduce(
			(total, spy) =>
				total +
				spy.mock.calls.filter(
					(args, index) =>
						(spy.mock.contexts[index] as XmlElement).localName === 'body' &&
						(spy !== moves[0] || args[1] !== null),
				).length,
			0,
		);
		vi.restoreAllMocks();
		expect(calls).toBeLessThan(100);
		const xml = await (await JSZip.loadAsync(saved)).file('word/document.xml')!.async('string');
		const reloaded = await loadDocx(saved);
		expect(reloaded.model.blocks).toHaveLength(3001);
		expect(xml.indexOf('First')).toBeLessThan(xml.indexOf('Paragraph 0<'));
		expect(xml.indexOf('Paragraph 2999<')).toBeGreaterThan(xml.indexOf('Paragraph 0<'));
	});
});
