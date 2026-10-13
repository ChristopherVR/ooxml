import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PresentationBuilder } from '../../core/builders/sdk/PresentationBuilder';
import { PptxHandler } from '../../core/PptxHandler';
import type { PptxSlide } from '../../core/types';
import { cloneSlide } from '../../core/utils/clone-utils';

/**
 * A rounded "pill" freeform interleaves its commands (moveTo, curve, curve,
 * line, curve, curve, line, close). The parsed `a:pathLst` object groups them
 * by type, and the source order rides on that exact object (a WeakMap). The
 * undo history, duplicate and paste all deep-clone the slide, and a clone
 * that rebuilt `customGeometryRawData.pathLstXml` lost the order: the save
 * re-derived the path from the clone and wrote every `moveTo` first, then the
 * curves, then the lines, so each pill came back as a bowtie.
 */
const PILL = [
	'<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/>',
	'<a:pathLst><a:path w="400" h="100">',
	'<a:moveTo><a:pt x="50" y="0"/></a:moveTo>',
	'<a:cubicBezTo><a:pt x="22" y="0"/><a:pt x="0" y="22"/><a:pt x="0" y="50"/></a:cubicBezTo>',
	'<a:cubicBezTo><a:pt x="0" y="78"/><a:pt x="22" y="100"/><a:pt x="50" y="100"/></a:cubicBezTo>',
	'<a:lnTo><a:pt x="350" y="100"/></a:lnTo>',
	'<a:cubicBezTo><a:pt x="378" y="100"/><a:pt x="400" y="78"/><a:pt x="400" y="50"/></a:cubicBezTo>',
	'<a:cubicBezTo><a:pt x="400" y="22"/><a:pt x="378" y="0"/><a:pt x="350" y="0"/></a:cubicBezTo>',
	'<a:lnTo><a:pt x="50" y="0"/></a:lnTo>',
	'<a:close/>',
	'</a:path></a:pathLst></a:custGeom>',
].join('');

const PILL_ORDER = [
	'moveTo',
	'cubicBezTo',
	'cubicBezTo',
	'lnTo',
	'cubicBezTo',
	'cubicBezTo',
	'lnTo',
	'close',
];

function asArrayBuffer(bytes: Uint8Array): ArrayBuffer {
	return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

async function pillDeck(): Promise<Uint8Array> {
	const created = await PresentationBuilder.create();
	const slide = created
		.createSlide('Blank')
		.addShape('rect', { x: 40, y: 40, width: 400, height: 100 })
		.build();
	const zip = await JSZip.loadAsync(await created.handler.save([slide]));
	const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
	const injected = xml.replace(
		/<a:prstGeom\b[^>]*>(?:[\s\S]*?)<\/a:prstGeom>|<a:prstGeom\b[^>]*\/>/u,
		PILL,
	);
	expect(injected).not.toBe(xml);
	zip.file('ppt/slides/slide1.xml', injected);
	return zip.generateAsync({ type: 'uint8array' });
}

async function savedCommandOrder(handler: PptxHandler, slides: PptxSlide[]): Promise<string[]> {
	const zip = await JSZip.loadAsync(await handler.save(slides));
	const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
	const path = /<a:path\b[^>]*>([\s\S]*?)<\/a:path>/u.exec(xml)?.[1] ?? '';
	return [...path.matchAll(/<a:(moveTo|lnTo|arcTo|quadBezTo|cubicBezTo|close)\b/gu)].map(
		(match) => match[1] as string,
	);
}

describe('custom geometry command order survives a deep clone', () => {
	it('saves a dirty pill in source order (control)', async () => {
		const handler = new PptxHandler();
		const data = await handler.load(asArrayBuffer(await pillDeck()));
		const slide = data.slides[0]!;
		slide.isDirty = true;
		expect(await savedCommandOrder(handler, data.slides)).toStrictEqual(PILL_ORDER);
	});

	it('saves a cloned pill (undo snapshot, duplicate, paste) in source order', async () => {
		const handler = new PptxHandler();
		const data = await handler.load(asArrayBuffer(await pillDeck()));
		// What undo restores: a deep clone of the slide taken before the edit.
		const restored = cloneSlide(data.slides[0]!);
		restored.isDirty = true;
		expect(await savedCommandOrder(handler, [restored])).toStrictEqual(PILL_ORDER);
	});

	it('keeps the order through a clone of a clone (redo after undo)', async () => {
		const handler = new PptxHandler();
		const data = await handler.load(asArrayBuffer(await pillDeck()));
		const twice = cloneSlide(cloneSlide(data.slides[0]!));
		twice.isDirty = true;
		expect(await savedCommandOrder(handler, [twice])).toStrictEqual(PILL_ORDER);
	});

	it('keeps the order through structuredClone and a JSON round trip', async () => {
		const handler = new PptxHandler();
		const data = await handler.load(asArrayBuffer(await pillDeck()));
		const structured = structuredClone(data.slides[0]!);
		structured.isDirty = true;
		expect(await savedCommandOrder(handler, [structured])).toStrictEqual(PILL_ORDER);
		const json = JSON.parse(JSON.stringify(data.slides[0])) as PptxSlide;
		json.isDirty = true;
		expect(await savedCommandOrder(handler, [json])).toStrictEqual(PILL_ORDER);
	});
});
