import { Schema } from 'prosemirror-model';
import { describe, expect, it } from 'vitest';

import { imageNodeSpec } from './inline-content-schema';

const schema = new Schema({
	nodes: {
		doc: { content: 'paragraph+' },
		paragraph: { content: 'inline*', group: 'block' },
		text: { group: 'inline' },
		image: imageNodeSpec,
	},
});

describe('image node placeholder', () => {
	it('writes only numeric sizes into its style, whatever a peer put in the attributes', () => {
		const node = schema.nodes.image!.create({
			unsupported: 'Text box',
			widthPx: '1px;background:url(https://example.com/x)',
			heightPx: 30,
		});
		const spec = imageNodeSpec.toDOM!(node) as [string, Record<string, string>];
		expect(spec[1].style).toBe('width:0px;height:30px');
	});
});
