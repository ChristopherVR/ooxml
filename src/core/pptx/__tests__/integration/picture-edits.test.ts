import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { PptxHandler } from '../../core/PptxHandler';
import type { PptxSlide } from '../../core/types';
import {
	clearPictureCrop,
	isDirectPicture,
	pictureBorderPatch,
	removePictureBackgroundPatch,
	replacePicturePatch,
	resetPicturePatch,
} from '../../core/utils/picture-edits';

const ORIGINAL =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';
const REPLACEMENT =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mNgAAIAAAUAAen63NgAAAAASUVORK5CYII=';

function pictureAt(slide: PptxSlide | undefined, index = 0) {
	const picture = slide?.elements.filter(isDirectPicture)[index];
	if (!picture) throw new Error('Missing picture');
	return picture;
}

async function loadDeck() {
	const { handler, createSlide } = await PptxHandler.create();
	const first = createSlide('Blank')
		.addImage(ORIGINAL, { x: 15, y: 20, width: 120, height: 80, altText: 'Example' })
		.addImage(ORIGINAL, { x: 190, y: 20, width: 120, height: 80 })
		.build();
	const picture = pictureAt(first);
	picture.cropLeft = 0.25;
	picture.imageEffects = { alphaModFix: 60 };
	const bytes = await handler.save([first]);
	const editor = new PptxHandler();
	const { slides } = await editor.load(bytes);
	return { editor, slides, picture: pictureAt(slides[0]) };
}

describe('picture commands round-trip', () => {
	it('changes a picture without overwriting shared media or changing its frame, crop or effects', async () => {
		const { editor, slides, picture } = await loadDeck();
		Object.assign(picture, replacePicturePatch(REPLACEMENT));
		const bytes = await editor.save(slides);
		const { slides: reloaded } = await new PptxHandler().load(bytes);
		expect(pictureAt(reloaded[0])).toMatchObject({
			x: 15,
			y: 20,
			width: 120,
			height: 80,
			altText: 'Example',
			cropLeft: 0.25,
			imageEffects: { alphaModFix: 60 },
		});
		expect(pictureAt(reloaded[0]).imagePath).not.toBe(pictureAt(reloaded[0], 1).imagePath);
		const archive = await JSZip.loadAsync(bytes);
		expect(archive.file(/ppt\/media\/[^/]+\.png$/u)).toHaveLength(2);
	});

	it('omits detached compressed media from the saved deck and preserves it for undo', async () => {
		const { editor, slides, picture } = await loadDeck();
		const slide = slides[0]!;
		const before = [...slide.elements];
		slide.elements = [
			{ ...picture, ...replacePicturePatch(REPLACEMENT), ...clearPictureCrop() } as typeof picture,
		];
		const bytes = await editor.save(slides);
		const archive = await JSZip.loadAsync(bytes);
		expect(archive.file(/ppt\/media\/[^/]+\.png$/u)).toHaveLength(1);
		const loaded = await new PptxHandler().load(bytes);
		expect(pictureAt(loaded.slides[0]).cropLeft ?? 0).toBe(0);
		slide.elements = before;
		const undoArchive = await JSZip.loadAsync(await editor.save(slides));
		expect(undoArchive.file(picture.imagePath!)).not.toBeNull();
	});

	it('keeps the previous bitmap when another package part still references it', async () => {
		const { editor, slides, picture } = await loadDeck();
		const bytes = await editor.save(slides);
		const input = await JSZip.loadAsync(bytes);
		const rootRelationships = input.file('_rels/.rels')!;
		const relationships = await rootRelationships.async('string');
		input.file(
			'_rels/.rels',
			relationships.replace(
				'</Relationships>',
				'<Relationship Id="rIdPicture" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="' +
					picture.imagePath +
					'"/></Relationships>',
			),
		);
		const handler = new PptxHandler();
		const reopened = await handler.load(await input.generateAsync({ type: 'uint8array' }));
		const existing = pictureAt(reopened.slides[0]);
		reopened.slides[0]!.elements = [Object.assign(existing, replacePicturePatch(REPLACEMENT))];
		const output = await JSZip.loadAsync(await handler.save(reopened.slides));
		expect(output.file(picture.imagePath!)).not.toBeNull();
	});

	it('writes the background-removal marquee and both mark types, with a reloadable original', async () => {
		const { editor, slides, picture } = await loadDeck();
		const removal = {
			top: 0.1,
			bottom: 0.9,
			left: 0.2,
			right: 0.8,
			foregroundMarks: [{ x1: 0.2, y1: 0.3, x2: 0.4, y2: 0.5 }],
			backgroundMarks: [{ x1: 0.0, y1: 0.1, x2: 0.1, y2: 0.2 }],
		};
		Object.assign(picture, removePictureBackgroundPatch(picture, REPLACEMENT, ORIGINAL, removal));
		const bytes = await editor.save(slides);
		const zip = await JSZip.loadAsync(bytes);
		const markup = await zip.file(slides[0]!.id)?.async('string');
		expect(markup).toContain('<a14:backgroundRemoval t="10000" b="90000" l="20000" r="80000">');
		expect(markup).toContain('<a14:foregroundMark x1="20000" y1="30000" x2="40000" y2="50000"');
		expect(markup).toContain('<a14:backgroundMark x1="0" y1="10000" x2="10000" y2="20000"');
		const reopened = new PptxHandler();
		const loaded = await reopened.load(bytes);
		const changed = pictureAt(loaded.slides[0]);
		expect(changed.imageEffects?.backgroundRemoval).toMatchObject(removal);
		expect(changed.originalPicturePath).toBeTruthy();
		const original = await zip.file(changed.originalPicturePath!)?.async('base64');
		expect('data:image/png;base64,' + original).toBe(ORIGINAL);
	});

	it('writes the picture outline and clears picture formatting on reset', async () => {
		const { editor, slides, picture } = await loadDeck();
		Object.assign(
			picture,
			pictureBorderPatch(picture, { color: '#aa0044', width: 4, dash: 'dash' }),
		);
		const first = await new PptxHandler().load(await editor.save(slides));
		expect(pictureAt(first.slides[0]).shapeStyle).toMatchObject({
			strokeWidth: 4,
			strokeDash: 'dash',
		});
		Object.assign(picture, resetPicturePatch(picture));
		const reset = await new PptxHandler().load(await editor.save(slides));
		expect(pictureAt(reset.slides[0])).toMatchObject({ x: 15, y: 20, width: 120, height: 80 });
		expect(pictureAt(reset.slides[0]).cropLeft ?? 0).toBe(0);
		expect(pictureAt(reset.slides[0]).imageEffects?.alphaModFix).toBeUndefined();
		expect(pictureAt(reset.slides[0]).shapeStyle?.strokeWidth ?? 0).toBe(0);
		Object.assign(picture, resetPicturePatch(picture, { size: { width: 1, height: 1 } }));
		const resized = await new PptxHandler().load(await editor.save(slides));
		expect(pictureAt(resized.slides[0])).toMatchObject({ x: 15, y: 20, width: 1, height: 1 });
	});
});
