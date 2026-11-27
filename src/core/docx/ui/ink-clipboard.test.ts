// @vitest-environment jsdom
import { DOMParser as PMDOMParser, DOMSerializer, Schema } from 'prosemirror-model';
import { describe, expect, it } from 'vitest';
import { imageNodeSpec } from './inline-content-schema';

const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: { content: 'inline*', toDOM: () => ['p', 0], parseDOM: [{ tag: 'p' }] },
		text: { group: 'inline' },
		image: { ...imageNodeSpec, group: 'inline', inline: true },
	},
});

describe('ink through the clipboard', () => {
	it('copies strokes, size and position, without the source part', () => {
		const ink = {
			relId: 'rId4',
			partName: 'word/ink/ink1.xml',
			strokes: [{ path: 'M 0 0 L 10 10', color: '#000000', width: 2, opacity: 1 }],
			width: 12,
			height: 12,
		};
		const node = schema.nodes.image!.create({
			unsupported: 'Ink',
			widthPx: 40,
			heightPx: 30,
			anchored: true,
			drawingName: 'Ink 1',
			placement: JSON.stringify({ wrap: 'none', offsetXPx: 5, offsetYPx: 6 }),
			ink: JSON.stringify(ink),
		});
		const dom = DOMSerializer.fromSchema(schema).serializeNode(node) as HTMLElement;
		const holder = document.createElement('p');
		holder.append(dom);
		const parsed = PMDOMParser.fromSchema(schema).parse(holder).firstChild!.firstChild!;
		expect(parsed.attrs.widthPx).toBe(40);
		expect(parsed.attrs.anchored).toBe(true);
		expect(JSON.parse(parsed.attrs.placement)).toMatchObject({ offsetXPx: 5 });
		const pasted = JSON.parse(parsed.attrs.ink);
		expect(pasted.strokes).toEqual(ink.strokes);
		expect(pasted.relId).toBeUndefined();
		expect(pasted.partName).toBeUndefined();
	});

	it('reads nothing from a placeholder whose ink is not ink', () => {
		const holder = document.createElement('p');
		holder.innerHTML =
			'<span data-docx-image-placeholder="1" data-ink=\'{"strokes":[{"path":1}],"width":1,"height":1}\'>Ink</span>';
		const parsed = PMDOMParser.fromSchema(schema).parse(holder).firstChild!.firstChild!;
		expect(parsed.attrs.ink).toBeNull();
	});
});
