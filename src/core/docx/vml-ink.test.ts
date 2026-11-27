import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { inkPaint } from './layout/ink-paint';
import { loadDocx, type DocumentModel } from './index';
import type { InlineImage } from './model';
import { parseVmlInk, vmlPathPolylines } from './vml-ink';
import { parseXml } from './xml';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const NS =
	`xmlns:w="${W}" xmlns:v="urn:schemas-microsoft-com:vml" ` +
	'xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w10="urn:schemas-microsoft-com:office:word"';

/** Legacy ink as Word 2003 writes it: the strokes as the shape's path and as ISF in `o:ink`. */
const SHAPE =
	'<v:shape id="_x0000_s1026" style="position:absolute;margin-left:36pt;margin-top:12pt;width:75pt;height:30pt;z-index:1"' +
	' coordorigin="100,100" coordsize="1000,400" path="m100,100l300,500,600,100e m700,200c800,300,900,300,1100,200e"' +
	' strokecolor="red" strokeweight="1.5pt" filled="f">' +
	'<v:stroke endcap="round"/><v:path shadowok="f" o:extrusionok="f" fillok="f" insetpenok="f"/>' +
	'<o:lock v:ext="edit" rotation="t" text="t"/><o:ink i="AHkdBIQBhgEBEFjPVIrml8VPjwb4utLhmyI=" annotation="t"/>' +
	'<w10:wrap type="none"/></v:shape>';

async function legacyDocument(): Promise<Uint8Array> {
	const source = readFileSync(
		path.join(import.meta.dirname, '__fixtures__/ink', 'word-ink-inline.docx'),
	);
	const zip = await JSZip.loadAsync(source);
	zip.file(
		'word/document.xml',
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document ${NS}><w:body>` +
			`<w:p><w:r><w:t>Signed:</w:t></w:r><w:r><w:pict>${SHAPE}</w:pict></w:r></w:p>` +
			'<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>',
	);
	return zip.generateAsync({ type: 'uint8array' });
}

function legacyImage(model: DocumentModel): InlineImage {
	for (const block of model.blocks)
		if (block.type === 'paragraph')
			for (const run of block.runs) if (run.image?.legacyInk) return run.image;
	throw new Error('no legacy ink');
}

describe('legacy VML ink', () => {
	it('reads the sub-paths of a VML path, sampling curves and reading empty numbers as 0', () => {
		const lines = vmlPathPolylines('m0,0l10,0,10,10x e m,5r5,0 c 5,5,10,5,10,0');
		expect(lines[0]).toEqual([
			{ x: 0, y: 0 },
			{ x: 10, y: 0 },
			{ x: 10, y: 10 },
			{ x: 0, y: 0 },
		]);
		expect(lines[1]?.[0]).toEqual({ x: 0, y: 5 });
		expect(lines[1]?.[1]).toEqual({ x: 5, y: 5 });
		expect(lines[1]?.at(-1)).toEqual({ x: 10, y: 0 });
		expect(lines[1]).toHaveLength(2 + 8);
		expect(vmlPathPolylines('')).toEqual([]);
		expect(vmlPathPolylines('m0,0 wa0,0,10,10,0,0,10,10')).toEqual([]);
	});

	it('maps the strokes from the coordinate space onto the box with the stroke colour and weight', () => {
		const shape = parseXml(`<v:shape ${NS}/>`).documentElement;
		expect(parseVmlInk(shape, 10, 10)).toBeUndefined();
		const ink = parseVmlInk(
			parseXml(SHAPE.replace('<v:shape ', `<v:shape ${NS} `)).documentElement,
			100,
			40,
		);
		expect(ink?.strokes).toHaveLength(2);
		expect(ink?.strokes[0]).toMatchObject({
			path: 'M 0 0 L 20 40 L 50 0',
			color: '#FF0000',
			width: 2,
			opacity: 1,
		});
	});

	it('draws legacy ink read only, floating where its style puts it, and saves the VML as read', async () => {
		const loaded = await loadDocx(await legacyDocument());
		const { model } = loaded;
		const image = legacyImage(model);
		expect(image).toMatchObject({
			unsupported: 'Legacy ink',
			anchored: true,
			vmlFloat: true,
			widthPx: 100,
			heightPx: 40,
			placement: { offsetXPx: 48, offsetYPx: 16 },
		});
		// Not the Draw tab's ink: painted, but never marked for its tools.
		expect(image.ink).toBeUndefined();
		const paint = inkPaint(image).chart;
		expect(paint?.svg).toContain('stroke="#FF0000"');
		expect(paint?.ink).toBeUndefined();
		expect(model.warnings.join(' ')).toContain('Legacy ink');
		const saved = await JSZip.loadAsync(await loaded.save(model));
		const xml = (await saved.file('word/document.xml')?.async('string')) ?? '';
		expect(xml).toContain('o:ink');
		expect(xml).toContain('AHkdBIQBhgEBEFjPVIrml8VPjwb4utLhmyI=');
		expect(xml).toContain('path="m100,100l300,500,600,100e');
	});

	it('moves floating legacy ink by patching its style', async () => {
		const loaded = await loadDocx(await legacyDocument());
		const { model } = loaded;
		const image = legacyImage(model);
		image.placement = { ...image.placement!, offsetXPx: 96 };
		const saved = await JSZip.loadAsync(await loaded.save(model));
		const xml = (await saved.file('word/document.xml')?.async('string')) ?? '';
		expect(xml).toContain('margin-left:72pt');
		expect(xml).toContain('o:ink');
	});
});
