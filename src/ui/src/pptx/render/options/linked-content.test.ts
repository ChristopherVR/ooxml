import { describe, expect, it } from 'vitest';

import {
	hasLinkedContent,
	isWebAddress,
	linkedContentUrl,
	linkedContentView,
	restoreLinkedContent,
	withholdLinkedContent,
} from './linked-content';

const WEB = 'https://cdn.test/picture.png';
const DATA = 'data:image/png;base64,AAAA';

/** One slide with every kind of linked content a deck can hold, and one embedded picture. */
function deck() {
	return [
		{
			id: 's1',
			backgroundImage: WEB,
			elements: [
				{ id: 'picture', type: 'picture', imagePath: WEB, imageData: WEB, altText: 'Logo' },
				{ id: 'embedded', type: 'picture', imagePath: 'ppt/media/a.png', imageData: DATA },
				{ id: 'svg', type: 'picture', svgPath: WEB, svgData: WEB },
				{
					id: 'video',
					type: 'media',
					mediaPath: 'https://www.youtube.com/watch?v=abc',
					posterFramePath: WEB,
					posterFrameData: WEB,
				},
				{ id: 'model', type: 'model3d', posterImage: WEB, imageData: WEB, imagePath: WEB },
				{ id: 'ole', type: 'ole', previewImageData: WEB },
				{
					id: 'shape',
					type: 'shape',
					shapeStyle: { fillMode: 'image', fillImageUrl: WEB },
					hyperlink: 'https://example.test/page',
					textSegments: [
						{
							text: 'One',
							bulletInfo: { imageRelId: 'rId4', imageDataUrl: WEB },
							style: { textFillBlipUrl: WEB },
						},
					],
				},
				{
					id: 'table',
					type: 'table',
					tableData: {
						rows: [{ cells: [{ style: { backgroundImageFillPath: WEB } }] }],
					},
				},
				{ id: 'group', type: 'group', children: [{ id: 'child', imageData: WEB }] },
			],
		},
	];
}

describe('withholdLinkedContent', () => {
	it('removes every linked web address a surface would fetch', () => {
		const view = withholdLinkedContent(deck());
		expect(hasLinkedContent(view)).toBe(false);
		expect(JSON.stringify(view)).not.toContain('cdn.test');
		expect(JSON.stringify(view)).not.toContain('youtube');
	});

	it('keeps embedded pictures, hyperlinks, text and ids', () => {
		const [slide] = withholdLinkedContent(deck());
		const byId = new Map(slide!.elements.map((element) => [element.id, element]));
		expect(byId.get('embedded')).toStrictEqual(deck()[0]!.elements[1]);
		expect(byId.get('picture')).toStrictEqual({ id: 'picture', type: 'picture', altText: 'Logo' });
		expect(withholdLinkedContent({ text: WEB, altText: WEB })).toStrictEqual({
			text: WEB,
			altText: WEB,
		});
		expect((byId.get('shape') as { hyperlink: string }).hyperlink).toBe(
			'https://example.test/page',
		);
		expect(
			(byId.get('shape') as { textSegments: { bulletInfo: object }[] }).textSegments[0]!.bulletInfo,
		).toStrictEqual({ imageRelId: 'rId4' });
	});

	it('withholds a table style picture fill', () => {
		const styles = { '{GUID}': { wholeTbl: { image: { path: WEB, data: WEB } } } };
		expect(withholdLinkedContent(styles)).toStrictEqual({ '{GUID}': { wholeTbl: { image: {} } } });
	});

	it('never changes the model, and returns unaffected objects as they are', () => {
		const model = deck();
		const before = JSON.stringify(model);
		const view = withholdLinkedContent(model);
		expect(JSON.stringify(model)).toBe(before);
		expect(view[0]!.elements[1]).toBe(model[0]!.elements[1]);
		expect(withholdLinkedContent(view)).toStrictEqual(view);
		expect(withholdLinkedContent(model)).toBe(view);
	});
});

describe('linkedContentView', () => {
	it('is the model itself once allowed', () => {
		const model = deck();
		expect(linkedContentView(model, true)).toBe(model);
		expect(hasLinkedContent(linkedContentView(model, false))).toBe(false);
	});
});

describe('hasLinkedContent', () => {
	it('finds linked content however deep it is', () => {
		expect(hasLinkedContent(deck())).toBe(true);
		expect(hasLinkedContent([{ id: 'g', children: [{ children: [{ imagePath: WEB }] }] }])).toBe(
			true,
		);
	});

	it('ignores embedded content and addresses that are only opened when clicked', () => {
		expect(
			hasLinkedContent([{ imageData: DATA, imagePath: 'ppt/media/a.png', hyperlink: WEB }]),
		).toBe(false);
		expect(hasLinkedContent(undefined)).toBe(false);
	});
});

describe('isWebAddress', () => {
	it('is true for http and https only', () => {
		expect(isWebAddress(WEB)).toBe(true);
		expect(isWebAddress('HTTP://cdn.test/a.png')).toBe(true);
		expect(isWebAddress(DATA)).toBe(false);
		expect(isWebAddress('blob:https://app.test/1')).toBe(false);
		expect(isWebAddress('ppt/media/https.png')).toBe(false);
	});
});

describe('linkedContentUrl', () => {
	it('withholds a web address until allowed, and passes the rest', () => {
		expect(linkedContentUrl(WEB, false)).toBeUndefined();
		expect(linkedContentUrl(WEB, true)).toBe(WEB);
		expect(linkedContentUrl(DATA, false)).toBe(DATA);
		expect(linkedContentUrl(undefined, false)).toBeUndefined();
	});
});

describe('restoreLinkedContent', () => {
	it('puts the links back into the parts an edit kept from what was drawn', () => {
		const model = {
			nodes: [{ id: 'a', text: 'A' }],
			shapes: [
				{ id: 's1', fillImageUrl: WEB },
				{ id: 's2', fillImageUrl: WEB },
			],
		};
		const drawn = withholdLinkedContent(model);
		// An edit made from the drawn copy: a node's text changes, and shape s2 gets a colour.
		const edited = {
			...drawn,
			nodes: [{ ...drawn.nodes[0]!, text: 'B' }],
			shapes: [drawn.shapes[0]!, { ...drawn.shapes[1]!, fillColor: '#f00' }],
		};
		const restored = restoreLinkedContent(edited);
		expect(restored.shapes[0]).toBe(model.shapes[0]);
		expect(restored.nodes[0]!.text).toBe('B');
		// The rebuilt shape is what the edit made of it.
		expect(restored.shapes[1]).toStrictEqual({ id: 's2', fillColor: '#f00' });
	});

	it('returns model data unchanged', () => {
		const model = deck();
		expect(restoreLinkedContent(model)).toBe(model);
	});
});
