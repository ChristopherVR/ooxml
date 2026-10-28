import { pictureEffectFilter } from '../src/picture-effects';
import { applyPictureHslPixels } from '../../core/drawingml/picture-hsl-pixels';
/** Generate repeatable native/shared picture comparison inputs. Run: bun src/ui/scripts/picture-native-fixture.ts <output-dir> */
import { mkdir, writeFile, readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import {
	pictureGridPng,
	pictureGridPixels,
	rgbaPng,
} from '../../core/diagram/test-support/picture-grid';
import JSZip from 'jszip';
import { JSDOM } from 'jsdom';
import { buildBlankPresentationArchive } from '../../core/pptx/core/builders/sdk/PresentationBuilder';
import { parseXml } from '../../core/xml';
import { parseDrawingPictureFill } from '../../core/drawingml/picture-fill';
import { picturePaint } from '../src/picture-paint';
const output = resolve(process.argv[2] ?? '.picture-native');
await mkdir(output, { recursive: true });
const png = pictureGridPng();
const imageUrl = `data:image/png;base64,${png.toString('base64')}`;
const a = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const baked = await JSZip.loadAsync(
	await readFile(
		new URL('../../core/drawingml/__fixtures__/native-artistic-fill.pptx', import.meta.url),
	),
);
const bakedPng = await baked.file('ppt/media/image1.png')!.async('uint8array');
const bakedOriginal = await baked.file('ppt/media/hdphoto1.wdp')!.async('uint8array');
const bakedSlide = await baked.file('ppt/slides/slide1.xml')!.async('string');
const bakedEffects = bakedSlide
	.match(/<a:extLst>[\s\S]*?<\/a:extLst>/)![0]
	.replace('rId3', 'rIdOriginal');
const cases = [
	{ id: 'plain', effects: '', layout: '<a:stretch><a:fillRect/></a:stretch>' },
	{
		id: 'gray-lum-alpha',
		effects: '<a:grayscl/><a:lum bright="15000" contrast="25000"/><a:alphaModFix amt="70000"/>',
		layout: '<a:stretch><a:fillRect/></a:stretch>',
	},
	{
		id: 'duotone',
		effects: '<a:duotone><a:srgbClr val="113355"/><a:srgbClr val="FFCC22"/></a:duotone>',
		layout: '<a:stretch><a:fillRect/></a:stretch>',
	},
	{
		id: 'alpha-bilevel',
		effects: '<a:alphaBiLevel thresh="60000"/>',
		layout: '<a:stretch><a:fillRect/></a:stretch>',
	},
	{
		id: 'bilevel',
		effects: '<a:biLevel thresh="50000"/>',
		layout: '<a:stretch><a:fillRect/></a:stretch>',
	},
	{
		id: 'tile-flip-offset',
		effects: '',
		crop: '<a:srcRect l="10000" r="10000"/>',
		layout: '<a:tile sx="150000" sy="150000" tx="95250" ty="-47625" algn="ctr" flip="xy"/>',
	},
	{
		id: 'fill-rect',
		effects: '',
		layout: '<a:stretch><a:fillRect l="15000" t="10000" r="25000" b="20000"/></a:stretch>',
	},
	{
		id: 'rot-with-shape-false',
		effects: '',
		attributes: 'rotWithShape="0"',
		rotation: 30,
		layout: '<a:tile sx="200000" sy="200000" algn="ctr"/>',
	},

	...[
		['clr-repl', '<a:clrRepl><a:srgbClr val="FF7700"/></a:clrRepl>'],
		[
			'clr-change',
			'<a:clrChange useA="0"><a:clrFrom><a:srgbClr val="FF0000"/></a:clrFrom><a:clrTo><a:srgbClr val="FFFF00"/></a:clrTo></a:clrChange>',
		],
		['hsl', '<a:hsl hue="1800000" sat="50000" lum="20000"/>'],
		['tint', '<a:tint hue="3600000" amt="30000"/>'],
		['blur', '<a:blur rad="2" grow="0"/>'],
		['blur-six', '<a:blur rad="6" grow="0"/>'],
		['hsl-sat-only', '<a:hsl hue="0" sat="50000" lum="100000"/>'],
		['hsl-lum-only', '<a:hsl hue="0" sat="100000" lum="50000"/>'],
		['hsl-identity', '<a:hsl hue="0" sat="100000" lum="100000"/>'],
		['alpha-inv', '<a:alphaInv/>'],
		['alpha-inv-colour', '<a:alphaInv><a:srgbClr val="FF7700"/></a:alphaInv>'],
		['alpha-ceiling', '<a:alphaCeiling/>'],
		['alpha-floor', '<a:alphaFloor/>'],
		['alpha-mod', '<a:alphaMod><a:cont><a:alphaModFix amt="50000"/></a:cont></a:alphaMod>'],
		['alpha-repl', '<a:alphaRepl a="40000"/>'],
	].map(([id, effects]) => ({
		id: id!,
		effects: effects!,
		layout: '<a:stretch><a:fillRect/></a:stretch>',
	})),
	{
		id: 'native-baked-pencil',
		effects: bakedEffects,
		layout: '<a:stretch><a:fillRect/></a:stretch>',
	},
];
const archive = await JSZip.loadAsync(
	await buildBlankPresentationArchive({ initialSlideCount: cases.length }),
);
archive.file('ppt/media/picture-grid.png', png);
archive.file('ppt/media/baked-pencil.png', bakedPng);
archive.file('ppt/media/baked-original.wdp', bakedOriginal);
let content = await archive.file('[Content_Types].xml')!.async('string');
if (!content.includes('Extension="png"'))
	content = content.replace(
		'</Types>',
		'<Default Extension="png" ContentType="image/png"/></Types>',
	);
content = content.replace(
	'</Types>',
	'<Default Extension="wdp" ContentType="image/vnd.ms-photo"/></Types>',
);
archive.file('[Content_Types].xml', content);
const manifest = [];
for (const [index, item] of cases.entries()) {
	const fill = `<a:blipFill xmlns:a="${a}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ${item.attributes ?? ''}><a:blip r:embed="rIdPicture">${item.effects}</a:blip>${item.crop ?? ''}${item.layout}</a:blipFill>`;
	const slide = `ppt/slides/slide${index + 1}.xml`;
	let xml = await archive.file(slide)!.async('string');
	const shape = `<p:sp><p:nvSpPr><p:cNvPr id="2" name="${item.id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm rot="${(item.rotation ?? 0) * 60000}"><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="1219200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${fill.replace('<a:blipFill', '<a:blipFill')}</p:spPr></p:sp>`;
	xml = xml.replace('</p:spTree>', `${shape}</p:spTree>`);
	archive.file(slide, xml);
	const rel = `ppt/slides/_rels/slide${index + 1}.xml.rels`;
	const relXml = await archive.file(rel)!.async('string');
	archive.file(
		rel,
		relXml.replace(
			'</Relationships>',
			'<Relationship Id="rIdPicture" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/' +
				(item.id === 'native-baked-pencil' ? 'baked-pencil.png' : 'picture-grid.png') +
				'"/><Relationship Id="rIdOriginal" Type="http://schemas.microsoft.com/office/2007/relationships/hdphoto" Target="../media/baked-original.wdp"/></Relationships>',
		),
	);
	const document = new JSDOM('<!doctype html>').window.document;
	const ns = 'http://www.w3.org/2000/svg';
	const svg = document.createElementNS(ns, 'svg');
	svg.setAttribute('xmlns', ns);
	svg.setAttribute('width', '192');
	svg.setAttribute('height', '128');
	svg.setAttribute('viewBox', '0 0 192 128');
	const defs = document.createElementNS(ns, 'defs');
	svg.append(defs);
	const parsed = parseDrawingPictureFill(parseXml(fill).documentElement);
	let pixelsUrl: string | undefined;
	if (
		parsed.effects?.some((e) => e.kind === 'standard' && (e.name === 'hsl' || e.name === 'tint'))
	) {
		const pixels = pictureGridPixels();
		for (const effect of parsed.effects) {
			if (effect.kind === 'standard') applyPictureHslPixels(pixels, effect);
		}
		pixelsUrl = `data:image/png;base64,${rgbaPng(48, 32, pixels).toString('base64')}`;
	}
	if (parsed.effects?.some((e) => e.kind === 'standard' && e.name === 'blur')) {
		const stage = document.createElementNS(ns, 'svg');
		stage.setAttribute('xmlns', ns);
		stage.setAttribute('width', '48');
		stage.setAttribute('height', '32');
		const d = document.createElementNS(ns, 'defs');
		stage.append(d);
		const fx = pictureEffectFilter(document, d, parsed.effects);
		const img = document.createElementNS(ns, 'image');
		img.setAttribute('href', imageUrl);
		img.setAttribute('width', '48');
		img.setAttribute('height', '32');
		if (fx.filter) img.setAttribute('filter', fx.filter);
		stage.append(img);
		pixelsUrl = `data:image/svg+xml,${encodeURIComponent(stage.outerHTML)}`;
	}
	const paint = picturePaint(
		document,
		defs,
		pixelsUrl ? { ...parsed, effects: [] } : parsed,
		{
			url:
				pixelsUrl ??
				(item.id === 'native-baked-pencil'
					? `data:image/png;base64,${Buffer.from(bakedPng).toString('base64')}`
					: imageUrl),
			width: 48,
			height: 32,
		},
		{ width: 192, height: 128, rotation: item.rotation ?? 0 },
	);
	const rect = document.createElementNS(ns, 'rect');
	rect.setAttribute('width', '192');
	rect.setAttribute('height', '128');
	rect.setAttribute('fill', paint.paint);
	if (item.rotation) rect.setAttribute('transform', `rotate(${item.rotation} 96 64)`);
	svg.append(rect);
	await writeFile(resolve(output, `${item.id}.svg`), svg.outerHTML);
	manifest.push({
		id: item.id,
		slide: index + 1,
		shape: 1,
		width: 192,
		height: 128,
		sharedSvg: `${item.id}.svg`,
		expect: { maximumError: 0.02, minimumOverlap: 0.8 },
		note: item.rotation
			? 'Shape rotation and inverse fill rotation are clipped to the declared frame.'
			: undefined,
	});
}
await writeFile(
	resolve(output, 'pictures.pptx'),
	await archive.generateAsync({ type: 'uint8array' }),
);
await writeFile(
	resolve(output, 'manifest.json'),
	JSON.stringify({ presentation: 'pictures.pptx', cases: manifest }, null, 2),
);
console.log(output);
