import { expect, it } from 'vitest';
import { encodePng } from '../../drawingml/raster-encode';
import { NS } from '../../xml/index';
import { createEditSession, createWorkbook, loadXlsx, pictureAnchorAt, saveXlsx } from '../index';

const position = pictureAnchorAt({ row: 1, col: 1 }, 120, 60);
const reopen = async (book: ReturnType<typeof createWorkbook>) => loadXlsx(await saveXlsx(book));

it('formats a native shape and keeps its formatting through edits, grouping and save', async () => {
	const book = createWorkbook();
	const edit = createEditSession(book);
	const index = edit.addShape(0, position, { kind: 'preset', preset: 'rect', name: 'Box' });
	expect(edit.formatDrawing(0, index, { kind: 'geometry', preset: 'hexagon' })).toBe(true);
	edit.formatDrawing(0, index, { kind: 'shapeStyle', block: 'theme', row: 1, column: 'accent2' });
	edit.formatDrawing(0, index, { kind: 'fill', color: 'FF0000' });
	edit.formatDrawing(0, index, { kind: 'outline', color: '0000FF' });
	edit.formatDrawing(0, index, { kind: 'shadow', enabled: true });
	edit.formatDrawing(0, index, { kind: 'altText', text: 'Red <box> & outline' });
	edit.formatDrawing(0, index, { kind: 'size', width: 150, height: 72 });
	expect(edit.undo()).toBe(true);
	expect(book.sheets[0]?.drawings[index]?.anchor.ext).toEqual(position.ext);
	expect(edit.redo()).toBe(true);

	const loaded = await reopen(book);
	expect(loaded.sheets[0]?.drawings[index]).toMatchObject({
		anchor: { ext: { cx: 1428750, cy: 685800 } },
		shape: {
			geometry: 'hexagon',
			frame: { width: 1428750, height: 685800 },
			fill: { kind: 'solid', color: { value: 'FF0000' } },
			line: { fill: { kind: 'solid', color: { value: '0000FF' } } },
			shadow: { kind: 'outerShadow' },
			description: 'Red <box> & outline',
		},
	});
	const after = createEditSession(loaded);
	after.formatDrawing(0, index, { kind: 'outline', color: null });
	after.rotateDrawing(0, index, 90);
	const second = after.addShape(0, position, { kind: 'preset', preset: 'ellipse', name: 'Oval' });
	expect(after.groupDrawings(0, [index, second], 'Group')).toBe(0);
	const grouped = await reopen(loaded);
	expect(grouped.sheets[0]?.drawings[0]).toMatchObject({
		shape: {
			children: [
				{
					geometry: 'hexagon',
					rotation: 90,
					line: { fill: { kind: 'none' } },
					shadow: { kind: 'outerShadow' },
				},
				{ geometry: 'ellipse' },
			],
		},
	});
});

it('keeps formatting after text box and WordArt edits on the same native shape', async () => {
	const book = createWorkbook();
	const edit = createEditSession(book);
	const box = edit.addShape(0, position, { kind: 'textBox', name: 'Label', text: 'Before' });
	edit.formatDrawing(0, box, { kind: 'fill', color: '123456' });
	edit.setTextBoxText(0, box, 'After');
	const art = edit.addWordArt(0, position, { text: 'WordArt', style: { fill: '000000' } });
	edit.formatDrawing(0, art, { kind: 'altText', text: 'WordArt caption' });
	edit.editWordArt(0, art, {
		text: 'WordArt',
		style: { fill: 'AABBCC', outline: 'FFFFFF', shadow: '222222' },
	});
	const loaded = await reopen(book);
	expect(loaded.sheets[0]?.drawings[box]).toMatchObject({
		shape: { text: { text: 'After' }, fill: { color: { value: '123456' } } },
	});
	expect(loaded.sheets[0]?.drawings[art]).toMatchObject({
		shape: {
			description: 'WordArt caption',
			wordArt: { fill: 'AABBCC', outline: 'FFFFFF', shadow: '222222' },
		},
	});
});

it('round-trips picture formatting, resets it and restores it with undo', async () => {
	const book = createWorkbook();
	const edit = createEditSession(book);
	const bytes = encodePng(1, 1, Uint8Array.of(255, 0, 0));
	edit.addImage(0, bytes, 'image/png', position, 'Logo');
	edit.editPictureEffects(0, 0, { brightness: 0.25, contrast: -0.15, grayscale: true });
	edit.setPictureCrop(0, 0, { cropLeft: 0.1 });
	edit.formatDrawing(0, 0, { kind: 'pictureStyle', style: 'simpleFrameWhite' });
	edit.formatDrawing(0, 0, { kind: 'altText', text: 'Picture <logo>' });
	edit.formatDrawing(0, 0, { kind: 'size', width: 90, height: 72 });
	const loaded = await reopen(book);
	const picture = () => loaded.sheets[0]?.drawings[0];
	expect(picture()).toMatchObject({
		kind: 'image',
		description: 'Picture <logo>',
		anchor: { ext: { cx: 857250, cy: 685800 } },
		crop: { cropLeft: 0.1 },
		effects: [{ kind: 'grayscale' }, { kind: 'luminance', brightness: 0.25, contrast: -0.15 }],
		line: { fill: { color: { value: 'FFFFFF' } } },
		shadow: { kind: 'outerShadow' },
	});
	const authored = picture();
	if (authored?.kind !== 'image' || !authored.sourceXml) throw new Error('Missing picture source');
	authored.sourceXml = authored.sourceXml
		.replace('prst="rect"', 'prst="ellipse"')
		.replace(
			'</xdr:spPr>',
			'</xdr:spPr><xdr:style xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef></xdr:style>',
		)
		.replace('</a:effectLst>', '<a:reflection/></a:effectLst>');
	const change = createEditSession(loaded);
	expect(change.formatDrawing(0, 0, { kind: 'resetPicture' })).toBe(true);
	const reset = await reopen(loaded);
	expect(reset.sheets[0]?.drawings[0]).toMatchObject({
		description: 'Picture <logo>',
	});
	expect(reset.sheets[0]?.drawings[0]).toHaveProperty('crop', {});
	expect(reset.sheets[0]?.drawings[0]).not.toHaveProperty('effects');
	expect(reset.sheets[0]?.drawings[0]).not.toHaveProperty('shadow');
	expect(reset.sheets[0]?.drawings[0]).not.toHaveProperty('line');
	expect(reset.sheets[0]?.drawings[0]).toHaveProperty(
		'sourceXml',
		expect.stringContaining('prst="ellipse"'),
	);
	expect(reset.sheets[0]?.drawings[0]).toHaveProperty(
		'sourceXml',
		expect.stringContaining('<a:lnRef idx="1">'),
	);
	expect(reset.sheets[0]?.drawings[0]).toHaveProperty(
		'sourceXml',
		expect.not.stringContaining('<a:effectLst'),
	);
	expect(change.undo()).toBe(true);
	expect(picture()).toHaveProperty('crop.cropLeft', 0.1);

	const replacement = encodePng(1, 1, Uint8Array.of(0, 255, 0));
	const original = picture();
	const originalPart = original?.kind === 'image' ? original.partName : undefined;
	expect(change.formatDrawing(0, 0, { kind: 'replacePicture', bytes: replacement })).toBe(true);
	const replaced = await reopen(loaded);
	const replacedImage = replaced.sheets[0]?.drawings[0];
	expect(replacedImage).toMatchObject({
		kind: 'image',
		description: 'Picture <logo>',
		crop: { cropLeft: 0.1 },
	});
	expect(
		replacedImage?.kind === 'image' && replaced.source?.parts.get(replacedImage.partName),
	).toEqual(replacement);
	expect(change.undo()).toBe(true);
	expect(picture()).toMatchObject({ partName: originalPart });
	expect(change.redo()).toBe(true);
	expect(picture()).toMatchObject({
		kind: 'image',
		partName: replacedImage?.kind === 'image' ? replacedImage.partName : '',
	});
	expect(() =>
		change.formatDrawing(0, 0, { kind: 'replacePicture', bytes: Uint8Array.of(0) }),
	).toThrow();
});

it('adds a shadow before existing reflection and preserves it when the shadow is removed', async () => {
	const book = createWorkbook();
	const bytes = encodePng(1, 1, Uint8Array.of(255, 0, 0));
	createEditSession(book).addImage(0, bytes, 'image/png', position, 'Logo');
	const loaded = await reopen(book);
	const picture = loaded.sheets[0]?.drawings[0];
	if (picture?.kind !== 'image' || !picture.sourceXml) throw new Error('Missing picture source');
	picture.sourceXml = picture.sourceXml.replace(
		'</xdr:spPr>',
		`<a:effectLst xmlns:a="${NS.a}"><a:reflection blurRad="12700"/></a:effectLst></xdr:spPr>`,
	);
	const edit = createEditSession(loaded);
	edit.formatDrawing(0, 0, { kind: 'shadow', enabled: true });
	const withShadow = (await reopen(loaded)).sheets[0]?.drawings[0];
	expect(withShadow).toHaveProperty('shadow.kind', 'outerShadow');
	expect(withShadow).toHaveProperty(
		'sourceXml',
		expect.stringMatching(/<a:outerShdw[\s\S]*<a:reflection/),
	);

	edit.formatDrawing(0, 0, { kind: 'shadow', enabled: false });
	const withoutShadow = (await reopen(loaded)).sheets[0]?.drawings[0];
	expect(withoutShadow).not.toHaveProperty('shadow');
	expect(withoutShadow).toHaveProperty(
		'sourceXml',
		expect.stringContaining('<a:reflection blurRad="12700"/>'),
	);
});
