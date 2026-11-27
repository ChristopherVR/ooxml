import { readFileSync } from 'node:fs';
import path from 'node:path';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { inkStrokeFromSamples } from '../drawingml/inkml/edit';
import { createDocument, loadDocx, saveDocx, type DocumentModel } from './index';
import { inksIn } from './ink';
import type { InlineImage, Paragraph } from './model';

const fixture = (name: string) =>
	new Uint8Array(readFileSync(path.join(import.meta.dirname, '__fixtures__/ink', name)));

/** The first ink drawing of the body. */
function inkImage(model: DocumentModel): InlineImage {
	for (const block of model.blocks)
		if (block.type === 'paragraph')
			for (const run of block.runs) if (run.image?.ink) return run.image;
	throw new Error('no ink');
}

describe('Word ink', () => {
	// Word 16's own save of a document holding PowerPoint ink (`generate-word-ink.ps1`).
	it('reads an inline ink drawing and its strokes', async () => {
		const { model } = await loadDocx(fixture('word-ink-inline.docx'));
		const image = inkImage(model);
		expect(image.unsupported).toBe('Ink');
		expect(image.anchored).toBeUndefined();
		expect(image.ink?.partName).toBe('word/ink/ink1.xml');
		expect(image.ink?.strokes).toHaveLength(4);
		expect(image.ink?.strokes[0]?.color).toBe('#E81123');
		expect(model.warnings.join(' ')).not.toContain('Ink content');
	});

	it('reads a floating ink drawing from inside its mc:AlternateContent', async () => {
		const { model } = await loadDocx(fixture('word-ink-floating.docx'));
		const image = inkImage(model);
		expect(image.anchored).toBe(true);
		expect(image.placement?.wrap).toBe('none');
		expect(image.ink?.strokes.length).toBeGreaterThan(0);
	});

	it('saves untouched ink byte for byte', async () => {
		const original = fixture('word-ink-inline.docx');
		const { model } = await loadDocx(original);
		const saved = await JSZip.loadAsync(await saveDocx(model));
		const source = await JSZip.loadAsync(original);
		expect(await saved.file('word/ink/ink1.xml')!.async('string')).toBe(
			await source.file('word/ink/ink1.xml')!.async('string'),
		);
		expect(await saved.file('word/document.xml')!.async('string')).toContain('w14:contentPart');
	});

	it('drops the stale fallback picture of ink whose strokes changed', async () => {
		const loaded = await loadDocx(fixture('word-ink-floating.docx'));
		const source = await JSZip.loadAsync(fixture('word-ink-floating.docx'));
		const fallbacks = (xml: string) => (xml.match(/<mc:Fallback/g) ?? []).length;
		const before = fallbacks(await source.file('word/document.xml')!.async('string'));
		const edited = structuredClone(loaded.model);
		inkImage(edited).ink!.strokes.splice(0, 1);
		const zip = await JSZip.loadAsync(await loaded.save(edited));
		const xml = await zip.file('word/document.xml')!.async('string');
		expect(fallbacks(xml)).toBeLessThan(before);
		expect(xml).toContain('<w14:contentPart');
	});

	it('saves erased strokes and reads them back', async () => {
		const loaded = await loadDocx(fixture('word-ink-floating.docx'));
		const edited = structuredClone(loaded.model);
		const ink = inkImage(edited).ink!;
		const kept = ink.strokes.slice(1);
		ink.strokes = kept;
		const reloaded = await loadDocx(await loaded.save(edited));
		const strokes = inkImage(reloaded.model).ink!.strokes;
		expect(strokes).toHaveLength(kept.length);
		expect(strokes.map((stroke) => stroke.path)).toEqual(kept.map((stroke) => stroke.path));
		expect(inkImage(reloaded.model).anchored).toBe(true);
	});

	it('writes new ink as a w14:contentPart with an InkML part Word opens', async () => {
		const model = createDocument();
		const stroke = inkStrokeFromSamples(
			[
				{ x: 2, y: 2, time: 0 },
				{ x: 40, y: 20, time: 16 },
				{ x: 80, y: 6, time: 32 },
			],
			{ kind: 'pen', color: '#0078D4', width: 2 },
		)!;
		const paragraph = model.blocks[0] as Paragraph;
		paragraph.runs.push({
			text: '',
			image: {
				relId: '',
				partName: '',
				contentType: 'application/octet-stream',
				widthPx: 82,
				heightPx: 22,
				anchored: true,
				placement: { wrap: 'none', offsetXPx: 30, offsetYPx: 12 },
				unsupported: 'Ink',
				ink: { strokes: [stroke], width: 82, height: 22 },
			},
		});
		const bytes = await saveDocx(model);
		const zip = await JSZip.loadAsync(bytes);
		const types = await zip.file('[Content_Types].xml')!.async('string');
		expect(types).toContain('PartName="/word/ink/ink1.xml" ContentType="application/inkml+xml"');
		const rels = await zip.file('word/_rels/document.xml.rels')!.async('string');
		expect(rels).toMatch(/relationships\/customXml" Target="ink\/ink1.xml"/);
		const document = await zip.file('word/document.xml')!.async('string');
		expect(document).toContain('wordprocessingInk');
		expect(document).toContain('<wp:anchor');
		const reloaded = await loadDocx(bytes);
		const image = inkImage(reloaded.model);
		expect(image.ink?.strokes[0]?.path).toBe(stroke.path);
		expect(image.ink?.strokes[0]?.color).toBe('#0078D4');
		expect(image.placement?.offsetXPx).toBe(30);
		expect(inksIn(reloaded.model.blocks)).toHaveLength(1);
	});

	it('stretches the strokes of a resized drawing onto its new box', async () => {
		const loaded = await loadDocx(fixture('word-ink-inline.docx'));
		const edited = structuredClone(loaded.model);
		const image = inkImage(edited);
		const before = image.ink!.strokes[1]!.path;
		image.widthPx *= 2;
		const reloaded = inkImage((await loadDocx(await loaded.save(edited))).model);
		expect(reloaded.widthPx).toBe(image.widthPx);
		const x = (path: string) => Number(/M\s*([\d.]+)/.exec(path)?.[1]);
		expect(x(reloaded.ink!.strokes[1]!.path)).toBeCloseTo(x(before) * 2, 0);
	});

	it('keeps loaded ink when new ink is drawn before it in its paragraph', async () => {
		const loaded = await loadDocx(fixture('word-ink-floating.docx'));
		const edited = structuredClone(loaded.model);
		const paragraph = edited.blocks.find(
			(block): block is Paragraph =>
				block.type === 'paragraph' && block.runs.some((run) => run.image?.ink),
		)!;
		const stroke = inkStrokeFromSamples(
			[
				{ x: 1, y: 1 },
				{ x: 40, y: 1 },
			],
			{ kind: 'pen', color: '#000000', width: 2 },
		)!;
		paragraph.runs.unshift({
			text: '',
			image: {
				relId: '',
				partName: '',
				contentType: 'application/octet-stream',
				widthPx: 42,
				heightPx: 3,
				anchored: true,
				placement: {
					wrap: 'none',
					relativeFrom: 'page',
					relativeFromV: 'page',
					offsetXPx: 50,
					offsetYPx: 60,
				},
				unsupported: 'Ink',
				ink: { strokes: [stroke], width: 42, height: 3 },
			},
		});
		const reloaded = await loadDocx(await loaded.save(edited));
		expect(inksIn(reloaded.model.blocks).map((ink) => ink.strokes.length)).toEqual([
			1,
			inkImage(loaded.model).ink!.strokes.length,
		]);
	});

	describe('ink moved, copied or deleted', () => {
		const inkParagraph = (model: DocumentModel) =>
			model.blocks.findIndex(
				(block) => block.type === 'paragraph' && block.runs.some((run) => run.image?.ink),
			);
		const inkRun = (model: DocumentModel) => {
			const paragraph = model.blocks[inkParagraph(model)] as Paragraph;
			return paragraph.runs.find((run) => run.image?.ink)!;
		};
		const parts = async (bytes: Uint8Array) =>
			Object.keys((await JSZip.loadAsync(bytes)).files).filter(
				(name) => name.startsWith('word/ink/') && !name.endsWith('/'),
			);
		const sourcePart = async () =>
			(await JSZip.loadAsync(fixture('word-ink-inline.docx')))
				.file('word/ink/ink1.xml')!
				.async('string');

		it('saves ink moved to another paragraph with its strokes and part', async () => {
			const loaded = await loadDocx(fixture('word-ink-inline.docx'));
			const edited = structuredClone(loaded.model);
			const run = inkRun(edited);
			const from = edited.blocks[inkParagraph(edited)] as Paragraph;
			from.runs.splice(from.runs.indexOf(run), 1);
			edited.blocks.push({ type: 'paragraph', runs: [{ text: 'moved ' }, run] } as Paragraph);
			const bytes = await loaded.save(edited);
			const reloaded = await loadDocx(bytes);
			const inks = inksIn(reloaded.model.blocks);
			expect(inks).toHaveLength(1);
			expect(inks[0]!.strokes).toHaveLength(4);
			expect(await parts(bytes)).toHaveLength(1);
			// Unchanged ink keeps Word's own bytes in its new part.
			const zip = await JSZip.loadAsync(bytes);
			expect(await zip.file(inks[0]!.partName!)!.async('string')).toBe(await sourcePart());
		});

		it('gives a copy of ink a part of its own', async () => {
			const loaded = await loadDocx(fixture('word-ink-inline.docx'));
			const edited = structuredClone(loaded.model);
			edited.blocks.push({
				type: 'paragraph',
				runs: [structuredClone(inkRun(edited))],
			} as Paragraph);
			const bytes = await loaded.save(edited);
			const inks = inksIn((await loadDocx(bytes)).model.blocks);
			expect(inks.map((ink) => ink.strokes.length)).toEqual([4, 4]);
			expect(new Set(inks.map((ink) => ink.partName)).size).toBe(2);
			expect(await parts(bytes)).toHaveLength(2);
			// Erasing in one copy leaves the other whole.
			const again = await loadDocx(bytes);
			const copy = structuredClone(again.model);
			inksIn(copy.blocks)[1]!.strokes.splice(0, 2);
			const after = inksIn((await loadDocx(await again.save(copy))).model.blocks);
			expect(after.map((ink) => ink.strokes.length)).toEqual([4, 2]);
		});

		it('keeps the original whole when an edited copy is saved', async () => {
			const loaded = await loadDocx(fixture('word-ink-inline.docx'));
			const edited = structuredClone(loaded.model);
			const copy = structuredClone(inkRun(edited));
			copy.image!.ink!.strokes.splice(0, 2);
			edited.blocks.push({ type: 'paragraph', runs: [copy] } as Paragraph);
			const bytes = await loaded.save(edited);
			const inks = inksIn((await loadDocx(bytes)).model.blocks);
			expect(inks.map((ink) => ink.strokes.length)).toEqual([4, 2]);
			const zip = await JSZip.loadAsync(bytes);
			expect(await zip.file('word/ink/ink1.xml')!.async('string')).toBe(await sourcePart());
		});

		/** A document whose default header holds a copy of the body's ink, saved and reloaded. */
		async function withHeaderInk() {
			const loaded = await loadDocx(fixture('word-ink-inline.docx'));
			const edited = structuredClone(loaded.model);
			const paragraph = structuredClone(edited.blocks[inkParagraph(edited)] as Paragraph);
			paragraph.id = 'header-ink';
			edited.sections![0]!.headers = {
				default: { blocks: [paragraph], partName: 'word/header1.xml' },
			};
			return loadDocx(await loaded.save(edited));
		}
		const headerInk = (model: DocumentModel, slot: 'default' | 'even') =>
			inksIn(model.sections![0]!.headers![slot]!.blocks)[0]!;

		it('gives a header copied from another its own ink parts', async () => {
			const loaded = await withHeaderInk();
			const source = headerInk(loaded.model, 'default');
			expect(source.strokes).toHaveLength(4);
			const edited = structuredClone(loaded.model);
			edited.sections![0]!.headers!.even = {
				...structuredClone(edited.sections![0]!.headers!.default!),
				partName: 'word/header2.xml',
				sourcePartName: 'word/header1.xml',
			};
			headerInk(edited, 'even').strokes.splice(0, 2);
			const reloaded = (await loadDocx(await loaded.save(edited))).model;
			expect(headerInk(reloaded, 'default').strokes).toHaveLength(4);
			expect(headerInk(reloaded, 'even').strokes).toHaveLength(2);
			expect(headerInk(reloaded, 'even').partName).not.toBe(
				headerInk(reloaded, 'default').partName,
			);
		});

		it('keeps header ink whole when an edited copy goes into the body', async () => {
			const loaded = await withHeaderInk();
			const edited = structuredClone(loaded.model);
			const run = structuredClone(
				(edited.sections![0]!.headers!.default!.blocks[0] as Paragraph).runs.find(
					(item) => item.image?.ink,
				)!,
			);
			run.image!.ink!.strokes.splice(0, 3);
			edited.blocks.push({ type: 'paragraph', runs: [run] } as Paragraph);
			const reloaded = (await loadDocx(await loaded.save(edited))).model;
			expect(headerInk(reloaded, 'default').strokes).toHaveLength(4);
			expect(inksIn(reloaded.blocks).at(-1)!.strokes).toHaveLength(1);
		});

		it('keeps ink through a paragraph split and join', async () => {
			const loaded = await loadDocx(fixture('word-ink-inline.docx'));
			const split = structuredClone(loaded.model);
			const at = inkParagraph(split);
			const paragraph = split.blocks[at] as Paragraph;
			const index = paragraph.runs.findIndex((run) => run.image?.ink);
			split.blocks.splice(at + 1, 0, {
				...paragraph,
				id: 'split-tail',
				runs: paragraph.runs.splice(index),
			} as Paragraph);
			const splitBytes = await loaded.save(split);
			const splitBack = await loadDocx(splitBytes);
			expect(inksIn(splitBack.model.blocks).map((ink) => ink.strokes.length)).toEqual([4]);
			const joined = structuredClone(splitBack.model);
			const next = joined.blocks.splice(at + 1, 1)[0] as Paragraph;
			(joined.blocks[at] as Paragraph).runs.push(...next.runs);
			const joinedBytes = await splitBack.save(joined);
			expect(inksIn((await loadDocx(joinedBytes)).model.blocks)).toHaveLength(1);
			expect(await parts(joinedBytes)).toHaveLength(1);
		});

		it('keeps ink when text is typed before it in its paragraph', async () => {
			const loaded = await loadDocx(fixture('word-ink-inline.docx'));
			const edited = structuredClone(loaded.model);
			(edited.blocks[inkParagraph(edited)] as Paragraph).runs.unshift({ text: 'New ' });
			const bytes = await loaded.save(edited);
			const zip = await JSZip.loadAsync(bytes);
			expect(await zip.file('word/ink/ink1.xml')!.async('string')).toBe(await sourcePart());
			expect(await zip.file('word/document.xml')!.async('string')).toContain('New ');
		});

		it('removes the part and relationship of deleted ink', async () => {
			const loaded = await loadDocx(fixture('word-ink-inline.docx'));
			const edited = structuredClone(loaded.model);
			const paragraph = edited.blocks[inkParagraph(edited)] as Paragraph;
			paragraph.runs = paragraph.runs.filter((run) => !run.image?.ink);
			const zip = await JSZip.loadAsync(await loaded.save(edited));
			expect(zip.file('word/ink/ink1.xml')).toBeNull();
			expect(await zip.file('word/_rels/document.xml.rels')!.async('string')).not.toContain(
				'ink/ink1.xml',
			);
			expect(await zip.file('[Content_Types].xml')!.async('string')).not.toContain('ink1.xml');
		});

		it('patches the box of resized ink and keeps its part', async () => {
			const loaded = await loadDocx(fixture('word-ink-inline.docx'));
			const edited = structuredClone(loaded.model);
			const image = inkRun(edited).image!;
			image.widthPx *= 2;
			const zip = await JSZip.loadAsync(await loaded.save(edited));
			const emu = Math.round(image.widthPx * 9525);
			expect(await zip.file('word/ink/ink1.xml')!.async('string')).toBe(await sourcePart());
			expect(await zip.file('word/document.xml')!.async('string')).toContain(`cx="${emu}"`);
		});

		it('never rewrites a part read only in part', async () => {
			const source = await JSZip.loadAsync(fixture('word-ink-inline.docx'));
			const big = Array.from({ length: 5000 }, (_, i) => `${i} ${i}`).join(',');
			const xml = (await sourcePart()).replace(
				'</inkml:ink>',
				`<inkml:trace>${big}</inkml:trace></inkml:ink>`,
			);
			expect(xml).toContain(big.slice(0, 20));
			source.file('word/ink/ink1.xml', xml);
			const loaded = await loadDocx(await source.generateAsync({ type: 'uint8array' }));
			const edited = structuredClone(loaded.model);
			const ink = inkRun(edited).image!.ink!;
			expect(ink.truncated).toBe(true);
			ink.strokes.splice(0, 1);
			edited.blocks.push({
				type: 'paragraph',
				runs: [structuredClone(inkRun(edited))],
			} as Paragraph);
			const zip = await JSZip.loadAsync(await loaded.save(edited));
			const names = await parts(await zip.generateAsync({ type: 'uint8array' }));
			expect(names).toHaveLength(2);
			for (const name of names) expect(await zip.file(name)!.async('string')).toBe(xml);
		});
	});
});
