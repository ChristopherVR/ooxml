import JSZip from 'jszip';
import type { PptxSlide } from 'pptx-viewer-core';
import { PptxHandler, PresentationBuilder } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { createInitialViewerState, createStore } from '../state';
import { createEditorOps } from './editor-operations';

// A rounded "pill" freeform: its commands interleave (moveTo, two curves, a
// line, two curves, a line, close). Undo restores a deep copy of the slide,
// and a copy used to lose that order, so the save wrote the commands grouped
// by type and every pill came back as a bowtie.
const PILL_GEOMETRY = [
	'<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/>',
	'<a:pathLst><a:path w="400" h="100">',
	'<a:moveTo><a:pt x="50" y="0"/></a:moveTo>',
	'<a:cubicBezTo><a:pt x="22" y="0"/><a:pt x="0" y="22"/><a:pt x="0" y="50"/></a:cubicBezTo>',
	'<a:cubicBezTo><a:pt x="0" y="78"/><a:pt x="22" y="100"/><a:pt x="50" y="100"/></a:cubicBezTo>',
	'<a:lnTo><a:pt x="350" y="100"/></a:lnTo>',
	'<a:cubicBezTo><a:pt x="378" y="100"/><a:pt x="400" y="78"/><a:pt x="400" y="50"/></a:cubicBezTo>',
	'<a:cubicBezTo><a:pt x="400" y="22"/><a:pt x="378" y="0"/><a:pt x="350" y="0"/></a:cubicBezTo>',
	'<a:lnTo><a:pt x="50" y="0"/></a:lnTo><a:close/>',
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

async function loadPillDeck(): Promise<{ handler: PptxHandler; slides: PptxSlide[] }> {
	const created = await PresentationBuilder.create();
	const built = created
		.createSlide('Blank')
		.addShape('rect', { x: 40, y: 40, width: 400, height: 100 })
		.build();
	const zip = await JSZip.loadAsync(await created.handler.save([built]));
	const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
	zip.file(
		'ppt/slides/slide1.xml',
		xml.replace(
			/<a:prstGeom\b[^>]*>(?:[\s\S]*?)<\/a:prstGeom>|<a:prstGeom\b[^>]*\/>/u,
			PILL_GEOMETRY,
		),
	);
	const bytes = await zip.generateAsync({ type: 'uint8array' });
	const handler = new PptxHandler();
	const data = await handler.load(bytes.slice().buffer);
	return { handler, slides: data.slides };
}

/** The `a:path` command names of slide 1 in a saved package. */
async function savedPillOrder(bytes: Uint8Array): Promise<string[]> {
	const zip = await JSZip.loadAsync(bytes);
	const xml = await zip.file('ppt/slides/slide1.xml')!.async('string');
	const path = /<a:path\b[^>]*>([\s\S]*?)<\/a:path>/u.exec(xml)?.[1] ?? '';
	return [...path.matchAll(/<a:(moveTo|lnTo|arcTo|quadBezTo|cubicBezTo|close)\b/gu)].map(
		(match) => match[1] ?? '',
	);
}

async function editorWithPill() {
	const { handler, slides } = await loadPillDeck();
	const store = createStore({
		...createInitialViewerState(),
		slides,
		currentSlide: 0,
		editable: true,
	});
	const ops = createEditorOps({ store, getHandler: () => handler, onHistoryChange: vi.fn() });
	return { ops };
}

describe('undo keeps custom geometry command order (vanilla)', () => {
	it('saves a pill in source order after an edit is undone', async () => {
		const { ops } = await editorWithPill();
		ops.commitNotes('an edit that is not a style edit');
		ops.undo();
		await expect(ops.save().then(savedPillOrder)).resolves.toStrictEqual(PILL_ORDER);
	});

	it('saves a pill in source order after undo then redo', async () => {
		const { ops } = await editorWithPill();
		ops.commitNotes('edited');
		ops.undo();
		ops.redo();
		await expect(ops.save().then(savedPillOrder)).resolves.toStrictEqual(PILL_ORDER);
	});
});
