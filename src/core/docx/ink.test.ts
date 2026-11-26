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
});
