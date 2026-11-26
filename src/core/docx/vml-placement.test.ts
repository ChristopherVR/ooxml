import { expect, it } from 'vitest';
import { parseDrawing, type DrawingContext } from './drawing';
import { buildXml, parseXml, REL_NS, WORD_NS } from './xml';
import { patchVmlFloat, vmlFloat, VML_NS } from './vml-placement';

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

const WORD10_NS = 'urn:schemas-microsoft-com:office:word';

it('patches only edited VML style declarations and adds missing wrapping', () => {
	const doc = parseXml(
		`<pict xmlns:v="${VML_NS}"><v:shape id="picture" style="position:absolute; WIDTH : 2in;custom: keep;z-index:-12;"><v:imagedata gain="19661f"/></v:shape></pict>`,
	);
	const pict = doc.documentElement;
	const shape = pict?.getElementsByTagNameNS(VML_NS, 'shape')[0];
	const data = shape?.getElementsByTagNameNS(VML_NS, 'imagedata')[0];
	if (!pict || !shape || !data) throw new Error('Expected a VML picture');
	const parsed = vmlFloat(shape, data);
	if (!parsed?.placement) throw new Error('Expected a VML float');
	const base = {
		relId: 'img1',
		partName: 'word/media/img.png',
		contentType: 'image/png',
		...parsed,
	};
	const before = buildXml(pict);
	patchVmlFloat(pict, structuredClone(base), base);
	expect(buildXml(pict)).toBe(before);

	patchVmlFloat(
		pict,
		{
			...base,
			widthPx: 0,
			placement: {
				...parsed.placement,
				wrap: 'tight',
				wrapText: 'bothSides',
				relativeHeight: 2 ** 31 - 7,
			},
		},
		base,
	);
	expect(shape.getAttribute('style')).toBe(
		'position:absolute; WIDTH :0pt;custom: keep;z-index:-7;',
	);
	expect(buildXml(shape)).toContain('<v:imagedata gain="19661f"/>');
	const wrap = shape.getElementsByTagNameNS(WORD10_NS, 'wrap')[0];
	expect(wrap?.getAttribute('type')).toBe('tight');
	expect(wrap?.getAttribute('side')).toBe('both');
	expect(vmlFloat(shape, data)?.placement).toMatchObject({
		relativeHeight: 2 ** 31 - 7,
		behindText: true,
	});
});
