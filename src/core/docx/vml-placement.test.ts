import { expect, it } from 'vitest';
import { parseDrawing, type DrawingContext } from './drawing';
import { parseXml, REL_NS, WORD_NS } from './xml';

const context: DrawingContext = {
	rels: new Map([
		['img', { id: 'img', type: `${REL_NS}/image`, target: 'media/image.png', mode: 'Internal' }],
	]),
	contentTypes: { defaults: new Map(), overrides: new Map() },
	mediaParts: new Set(['word/media/image.png']),
};

function picture(style: string, wrap = '', imageAttributes = '', relId = 'img') {
	const { documentElement } = parseXml(
		`<w:pict xmlns:w="${WORD_NS}" xmlns:r="${REL_NS}" xmlns:v="urn:schemas-microsoft-com:vml"
			xmlns:w10="urn:schemas-microsoft-com:office:word">
			<v:shape style="${style}"><v:imagedata r:id="${relId}" ${imageAttributes}/>${wrap}</v:shape>
		</w:pict>`,
	);
	return parseDrawing(documentElement, context);
}

it('reads a floated VML picture with alignment, real dimensions and washout', () => {
	const image = picture(
		'POSITION: Absolute; width:2.54cm; height:25.4mm; margin-left:10pt; margin-top:20pt;' +
			'mso-position-horizontal:right; mso-position-horizontal-relative:left-margin;' +
			'mso-position-vertical:bottom; mso-position-vertical-relative:top-margin; z-index:-5',
		'<w10:wrap type="square" side="both"/>',
		'gain="19661f" blacklevel="22938f"',
	);
	expect(image).toMatchObject({
		anchored: true,
		placement: {
			align: 'right',
			relativeFrom: 'leftMargin',
			alignV: 'bottom',
			relativeFromV: 'topMargin',
			wrap: 'square',
			wrapText: 'bothSides',
			behindText: true,
		},
		pictureFill: { kind: 'picture', effects: [{ kind: 'luminance' }] },
	});
	expect(image.widthPx).toBeCloseTo(96);
	expect(image.heightPx).toBeCloseTo(96);
	expect(image.placement?.offsetXPx).toBeUndefined();
	expect(image.placement?.offsetYPx).toBeUndefined();
});

it.each(['tight', 'through', 'topAndBottom', 'none'])(
	'reads %s wrapping and default anchors',
	(wrap) => {
		const image = picture(
			'position:absolute;width:20px;height:1in;margin-left:-.6pt;margin-top:0;z-index:3',
			`<w10:wrap type="${wrap}"/>`,
		);
		expect(image).toMatchObject({
			anchored: true,
			widthPx: 20,
			heightPx: 96,
			placement: { wrap, relativeFrom: 'column', relativeFromV: 'paragraph', offsetYPx: 0 },
		});
		expect(image.placement?.offsetXPx).toBeCloseTo(-0.8);
		expect(image.placement?.behindText).toBeUndefined();
		expect(image.pictureFill).toBeUndefined();
	},
);

it('keeps relative positioning inline and unresolved pictures as placeholders', () => {
	expect(picture('position:relative;width:12pt;height:9pt')).toMatchObject({
		widthPx: 16,
		heightPx: 12,
	});
	expect(picture('position:relative').anchored).toBeUndefined();
	expect(picture('position:absolute', '', '', 'missing')).toMatchObject({
		unsupported: 'Legacy VML drawing',
	});
});
