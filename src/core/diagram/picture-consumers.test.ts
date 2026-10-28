import { expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { withDiagramPicture } from './test-support/picture-drawing';
import { loadDocx } from '../docx/index';
import { loadXlsx } from '../xlsx/read/index';
import { saveXlsx } from '../xlsx/write/index';
import { createEditSession } from '../xlsx/edit/session';
it('loads Word cached pictures with internal media identities and preserves source drawing on save', async () => {
	const bytes = await withDiagramPicture(
		readFileSync(new URL('../docx/__fixtures__/smartart.docx', import.meta.url)),
		'word',
	);
	const loaded = await loadDocx(bytes);
	const image = loaded.model.blocks.flatMap((block) =>
		block.type === 'paragraph'
			? block.runs.flatMap((run) => (run.image?.diagram ? [run.image] : []))
			: [],
	)[0]!;
	const drawing = image.diagram!.drawing!;
	expect(loaded.media.get('word/media/shared-picture.png')?.length).toBeGreaterThan(0);
	expect(drawing.pictureParts).toEqual({ picture: 'word/media/shared-picture.png' });
	expect(drawing.shapes[0]).toMatchObject({
		geometry: 'ellipse',
		fill: {
			kind: 'picture',
			relId: 'picture',
			crop: { cropLeft: 0.25, cropRight: 0.25 },
			effects: [{ kind: 'grayscale' }, { kind: 'alphaModFix', amount: 0.5 }],
		},
	});
	const saved = await JSZip.loadAsync(await loaded.save(loaded.model));
	const before = await JSZip.loadAsync(bytes);
	expect(await saved.file('word/diagrams/drawing1.xml')!.async('string')).toBe(
		await before.file('word/diagrams/drawing1.xml')!.async('string'),
	);
});
it('loads Excel cached pictures and retains drawing, relationships and bytes through a move', async () => {
	const bytes = await withDiagramPicture(
		readFileSync(new URL('../xlsx/__fixtures__/excel-smartart.xlsx', import.meta.url)),
		'xl',
	);
	const wb = await loadXlsx(bytes);
	const index = wb.sheets[0]!.drawings.findIndex((drawing) => drawing.kind === 'smartArt');
	const object = wb.sheets[0]!.drawings[index]!;
	if (object.kind !== 'smartArt') throw new Error('Expected diagram');
	expect(object.diagram?.pictureParts).toEqual({ picture: 'xl/media/shared-picture.png' });
	expect(object.diagram?.shapes[0]?.fill?.kind).toBe('picture');
	createEditSession(wb).setDrawingAnchor(0, index, {
		from: { row: 5, col: 2, rowOffset: 0, colOffset: 0 },
		ext: { cx: 952500, cy: 476250 },
	});
	const saved = await saveXlsx(wb);
	const reread = await loadXlsx(saved);
	const again = reread.sheets[0]!.drawings[index]!;
	expect(again.kind === 'smartArt' && again.diagram?.pictureParts).toEqual({
		picture: 'xl/media/shared-picture.png',
	});
	const zip = await JSZip.loadAsync(saved);
	expect(zip.file('xl/media/shared-picture.png')).not.toBeNull();
	expect(await zip.file('xl/diagrams/_rels/drawing1.xml.rels')!.async('string')).toContain(
		'TargetMode="External"',
	);
});
