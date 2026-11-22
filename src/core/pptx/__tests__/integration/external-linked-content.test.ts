/**
 * `allowExternalImages` must cover every place a deck can link a picture from
 * another server, not only `p:pic`: a shape picture fill, a picture bullet and
 * a 3D model poster each resolve their own relationship, so each is checked
 * here against a deck whose image relationships all point at
 * {@link LINKED_ORIGIN} (`TargetMode="External"`).
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { PresentationBuilder } from '../../core/builders/sdk/PresentationBuilder';
import { PptxHandler } from '../../core/PptxHandler';
import type { Model3DPptxElement, PptxElement, PptxSlide } from '../../core/types';

const LINKED_ORIGIN = 'https://linked-content.test';
const IMAGE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const GLB_DATA = 'data:model/gltf-binary;base64,Z2xURgIAAAAMAAAA';
const PNG_DATA =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=';

const LINKED_FILL_SHAPE = `<p:sp><p:nvSpPr><p:cNvPr id="901" name="Linked fill"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:blipFill><a:blip r:link="rIdLinkedFill"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr></p:sp>`;
const LINKED_BULLET_SHAPE = `<p:sp><p:nvSpPr><p:cNvPr id="902" name="Linked bullet"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="1371600"/><a:ext cx="3657600" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr marL="342900" indent="-342900"><a:buBlip><a:blip r:embed="rIdLinkedBullet"/></a:buBlip></a:pPr><a:r><a:rPr lang="en-US"/><a:t>Linked bullet</a:t></a:r></a:p></p:txBody></p:sp>`;

/** A deck with a 3D model, a picture-filled shape and a picture bullet, all linked to another origin. */
async function linkedContentDeck(): Promise<ArrayBuffer> {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	const slide = createSlide('Blank').build();
	slide.elements.push({
		id: 'linked-model',
		type: 'model3d',
		x: 400,
		y: 80,
		width: 320,
		height: 180,
		modelData: GLB_DATA,
		modelMimeType: 'model/gltf-binary',
		posterImage: PNG_DATA,
	} satisfies Model3DPptxElement);
	data.slides.push(slide);
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	const slidePath = 'ppt/slides/slide1.xml';
	const relsPath = 'ppt/slides/_rels/slide1.xml.rels';
	const slideXml = await zip.file(slidePath)!.async('string');
	expect(slideXml).toContain('<am3d:blip r:embed="');
	zip.file(
		slidePath,
		slideXml.replace('</p:spTree>', `${LINKED_FILL_SHAPE}${LINKED_BULLET_SHAPE}</p:spTree>`),
	);
	const rels = await zip.file(relsPath)!.async('string');
	// Link the poster (and its fallback picture) instead of embedding it.
	const linkedRels = rels.replace(
		/(<Relationship\b[^>]*?)Target="\.\.\/media\/([^"]+\.png)"/gu,
		(_match, before: string, file: string) =>
			`${before}Target="${LINKED_ORIGIN}/poster-${file}" TargetMode="External"`,
	);
	expect(linkedRels).not.toBe(rels);
	zip.file(
		relsPath,
		linkedRels.replace(
			'</Relationships>',
			`<Relationship Id="rIdLinkedFill" Type="${IMAGE_REL}" Target="${LINKED_ORIGIN}/fill.png" TargetMode="External"/>` +
				`<Relationship Id="rIdLinkedBullet" Type="${IMAGE_REL}" Target="${LINKED_ORIGIN}/bullet.png" TargetMode="External"/></Relationships>`,
		),
	);
	const bytes = await zip.generateAsync({ type: 'uint8array' });
	return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

async function loadSlides(allowExternalImages: boolean, bytes?: ArrayBuffer): Promise<PptxSlide[]> {
	const handler = new PptxHandler();
	try {
		const data = await handler.load(bytes ?? (await linkedContentDeck()), {
			allowExternalImages,
			eagerDecodeImages: true,
		});
		return data.slides;
	} finally {
		handler.dispose();
	}
}

function byName(slide: PptxSlide, name: string): PptxElement {
	const element = slide.elements.find(
		(candidate) => (candidate as { name?: string }).name === name,
	);
	expect(element, name).toBeDefined();
	return element!;
}

describe('linked content beyond p:pic', () => {
	it('references no linked URL anywhere when allowExternalImages is false', async () => {
		const [slide] = await loadSlides(false);
		// Each element still loads; only its linked picture is dropped.
		expect(JSON.stringify(byName(slide, 'Linked fill'))).not.toContain(LINKED_ORIGIN);
		expect(JSON.stringify(byName(slide, 'Linked bullet'))).not.toContain(LINKED_ORIGIN);
		const model = slide.elements.find((element) => element.type === 'model3d');
		expect(model).toBeDefined();
		expect(JSON.stringify(model)).not.toContain(LINKED_ORIGIN);
		expect(JSON.stringify(slide)).not.toContain(LINKED_ORIGIN);
	});

	it('keeps each linked URL when allowExternalImages is true', async () => {
		const [slide] = await loadSlides(true);
		expect(byName(slide, 'Linked fill')).toMatchObject({
			imageData: `${LINKED_ORIGIN}/fill.png`,
		});
		const bullet = byName(slide, 'Linked bullet') as {
			textSegments?: { bulletInfo?: { imageDataUrl?: string } }[];
		};
		expect(bullet.textSegments?.some((segment) => segment.bulletInfo?.imageDataUrl)).toBe(true);
		expect(
			bullet.textSegments?.find((segment) => segment.bulletInfo?.imageDataUrl)?.bulletInfo
				?.imageDataUrl,
		).toBe(`${LINKED_ORIGIN}/bullet.png`);
		const model = slide.elements.find(
			(element): element is Model3DPptxElement => element.type === 'model3d',
		);
		expect(model?.posterImage).toMatch(/^https:\/\/linked-content\.test\/poster-/u);
	});
});

describe('a viewer model with its links', () => {
	it('saves an edited deck with its pictures still linked, never embedded', async () => {
		const handler = new PptxHandler();
		let saved: Uint8Array;
		try {
			// Viewers load with the addresses and decide per viewer what to fetch.
			const data = await handler.load(await linkedContentDeck(), { allowExternalImages: true });
			const slides = data.slides.map((slide) => ({
				...slide,
				isDirty: true,
				elements: slide.elements.map((element) =>
					(element as { name?: string }).name === 'Linked fill'
						? { ...element, x: element.x + 10 }
						: element,
				),
			}));
			saved = await handler.save(slides);
		} finally {
			handler.dispose();
		}
		const zip = await JSZip.loadAsync(saved);
		const slideXml = await zip.file('ppt/slides/slide1.xml')!.async('string');
		const rels = await zip.file('ppt/slides/_rels/slide1.xml.rels')!.async('string');
		expect(slideXml).toContain('r:link="rIdLinkedFill"');
		expect(rels).toContain(`Target="${LINKED_ORIGIN}/fill.png" TargetMode="External"`);
		expect(rels).toContain(`Target="${LINKED_ORIGIN}/bullet.png" TargetMode="External"`);
		expect(rels).toMatch(
			/Target="https:\/\/linked-content\.test\/poster-[^"]+" TargetMode="External"/u,
		);
		expect(Object.keys(zip.files).some((name) => name.includes('fill.png'))).toBe(false);

		const buffer = saved.buffer.slice(saved.byteOffset, saved.byteOffset + saved.byteLength);
		const [reopened] = await loadSlides(true, buffer as ArrayBuffer);
		expect(byName(reopened!, 'Linked fill')).toMatchObject({
			imageData: `${LINKED_ORIGIN}/fill.png`,
		});
	});
});

const WAV_DATA =
	'data:audio/wav;base64,UklGRiQAAABXQVZFZm10IBAAAAABAAEAQB8AAEAfAAABAAgAZGF0YQAAAAA=';
const AUDIO_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio';
const ACTION_SOUND_SHAPE = `<p:sp><p:nvSpPr><p:cNvPr id="903" name="Action sound"><a:hlinkClick r:id="" action="ppaction://noaction"><a:snd r:embed="rIdLinkedClickSound" name="click.wav"/></a:hlinkClick></p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>`;

/** A slide whose transition sound, effect sound and hyperlink sound are linked to {@link LINKED_ORIGIN}. */
async function linkedSoundDeck(): Promise<ArrayBuffer> {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	const slide = createSlide('Blank').build();
	slide.elements.push({
		id: 'animated',
		type: 'shape',
		x: 0,
		y: 0,
		width: 100,
		height: 100,
		shapeType: 'rect',
	} as PptxElement);
	slide.transition = { type: 'fade', durationMs: 300, soundData: WAV_DATA } as never;
	slide.animations = [
		{
			id: 'a1',
			elementId: 'animated',
			entrance: 'appear',
			trigger: 'onClick',
			soundData: WAV_DATA,
		},
	] as never;
	data.slides.push(slide);
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	const slidePath = 'ppt/slides/slide1.xml';
	const relsPath = 'ppt/slides/_rels/slide1.xml.rels';
	const slideXml = await zip.file(slidePath)!.async('string');
	zip.file(slidePath, slideXml.replace('</p:spTree>', `${ACTION_SOUND_SHAPE}</p:spTree>`));
	const rels = await zip.file(relsPath)!.async('string');
	const linkedRels = rels.replace(
		/(<Relationship\b[^>]*?)Target="\.\.\/media\/([^"]+\.wav)"/gu,
		(_match, before: string, file: string) =>
			`${before}Target="${LINKED_ORIGIN}/${file}" TargetMode="External"`,
	);
	expect(linkedRels.match(/TargetMode="External"/gu)?.length).toBeGreaterThanOrEqual(2);
	zip.file(
		relsPath,
		linkedRels.replace(
			'</Relationships>',
			`<Relationship Id="rIdLinkedClickSound" Type="${AUDIO_REL}" Target="${LINKED_ORIGIN}/click.wav" TargetMode="External"/></Relationships>`,
		),
	);
	const bytes = await zip.generateAsync({ type: 'uint8array' });
	return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

describe('linked sounds', () => {
	it('keeps the linked address of a transition, effect and hyperlink sound', async () => {
		const [slide] = await loadSlides(true, await linkedSoundDeck());
		const linkedSound = /^https:\/\/linked-content\.test\/[^/]+\.wav$/u;
		expect(slide!.transition?.soundPath).toMatch(linkedSound);
		expect(slide!.nativeAnimations?.find((anim) => anim.soundPath)?.soundPath).toMatch(linkedSound);
		expect(byName(slide!, 'Action sound').actionClick?.soundPath).toBe(
			`${LINKED_ORIGIN}/click.wav`,
		);
	});

	it('references none of them when allowExternalImages is false', async () => {
		const [slide] = await loadSlides(false, await linkedSoundDeck());
		expect(slide!.transition?.soundRId).toBeDefined();
		expect(slide!.transition?.soundPath).toBeUndefined();
		expect(JSON.stringify(slide)).not.toContain(LINKED_ORIGIN);
	});
});
