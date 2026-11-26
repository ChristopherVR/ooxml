import { expect, it } from 'vitest';
import { PICTURE_STYLES } from '../../drawingml/picture-styles-catalog';
import { encodePng } from '../../drawingml/raster-encode';
import { NS, elements, first, parseXml } from '../../xml/index';
import {
	createEditSession,
	createWorkbook,
	loadXlsx,
	pictureAnchorAt,
	saveXlsx,
	shapeSvg,
} from '../index';

const anchor = pictureAnchorAt({ row: 1, col: 1 }, 120, 60);

it('writes bare theme references and keeps geometry and text, then writes an explicit preset', async () => {
	const book = createWorkbook();
	const edit = createEditSession(book);
	const index = edit.addShape(0, anchor, { kind: 'textBox', text: 'Keep me', name: 'Label' });
	edit.formatDrawing(0, index, { kind: 'outline', color: 'FF0000' });
	edit.formatDrawing(0, index, { kind: 'shadow', enabled: true });
	edit.formatDrawing(0, index, { kind: 'shapeStyle', block: 'theme', row: 5, column: 'accent2' });
	const loaded = await loadXlsx(await saveXlsx(book));
	const drawing = loaded.sheets[0]?.drawings[index];
	if (drawing?.kind !== 'unsupported' || !drawing.shape || !drawing.sourceXml)
		throw new Error('Missing shape');
	const root = parseXml(drawing.sourceXml).documentElement;
	const shape = first(root, 'sp', NS.xdr);
	const props = first(shape, 'spPr', NS.xdr);
	const style = first(shape, 'style', NS.xdr);
	if (!props || !style) throw new Error('Missing shape formatting');
	expect(elements(props).map(({ localName }) => localName)).toEqual(['xfrm', 'prstGeom']);
	const refs = elements(style);
	expect(refs.map((ref) => [ref.localName, ref.getAttribute('idx')])).toEqual([
		['lnRef', '0'],
		['fillRef', '3'],
		['effectRef', '3'],
		['fontRef', 'minor'],
	]);
	expect(drawing.shape).toMatchObject({
		text: { text: 'Keep me' },
		style: {
			fill: { color: { value: 'accent2' } },
			effect: { color: { value: 'accent2' } },
			font: { color: { value: 'lt1' } },
		},
	});
	const svg = shapeSvg(drawing.shape, 120, 60, { theme: loaded.theme });
	expect(svg).toContain('<linearGradient');
	expect(svg).toContain('<feDropShadow');

	const reedit = createEditSession(loaded);
	reedit.formatDrawing(0, index, {
		kind: 'shapeStyle',
		block: 'preset',
		row: 4,
		column: 'accent3',
	});
	const preset = (await loadXlsx(await saveXlsx(loaded))).sheets[0]?.drawings[index];
	expect(preset).toMatchObject({
		shape: {
			fill: {
				kind: 'gradient',
				stops: [{ position: 0, color: { value: 'accent3' } }, { position: 48 }, { position: 100 }],
			},
			line: { fill: { kind: 'none' } },
			style: { line: { idx: 0 }, fill: { idx: 0 }, effect: { idx: 0 } },
		},
	});
});

it('writes all 28 picture frames, including guards against stale effects and 3-D', async () => {
	const book = createWorkbook();
	const edit = createEditSession(book);
	const bytes = encodePng(1, 1, Uint8Array.of(255, 0, 0));
	for (const [index, spec] of PICTURE_STYLES.entries()) {
		edit.addImage(0, bytes, 'image/png', anchor, spec.label);
		edit.formatDrawing(0, index, { kind: 'pictureStyle', style: 'simpleFrameWhite' });
		edit.formatDrawing(0, index, { kind: 'pictureStyle', style: spec.key });
	}
	const pictures = (await loadXlsx(await saveXlsx(book))).sheets[0]?.drawings;
	expect(pictures).toHaveLength(28);
	for (const [index, spec] of PICTURE_STYLES.entries()) {
		const picture = pictures?.[index];
		if (picture?.kind !== 'image' || !picture.sourceXml) throw new Error('Missing picture');
		expect(picture.geometry).toBe(spec.geom);
		expect(picture.line?.widthEmu).toBe(spec.line?.w);
		const root = parseXml(picture.sourceXml).documentElement;
		const props = first(first(root, 'pic', NS.xdr), 'spPr', NS.xdr);
		expect(Boolean(first(props, 'scene3d', NS.a))).toBe(Boolean(spec.scene));
		expect(Boolean(first(props, 'sp3d', NS.a))).toBe(Boolean(spec.sp3d));
		const list = first(props, 'effectLst', NS.a);
		if (!list) throw new Error('Missing picture effect list');
		const effects = elements(list).map(({ localName }) => localName);
		expect(effects).toEqual([
			...(spec.innerBlur === undefined ? [] : ['innerShdw']),
			...(spec.outer ? ['outerShdw'] : []),
			...(spec.reflection ? ['reflection'] : []),
			...(spec.softEdge === undefined ? [] : ['softEdge']),
		]);
	}
});
