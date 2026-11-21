/**
 * The gate over a whole loaded deck, not a hand-built model: every web
 * address the parser puts in a viewer's model from a linked relationship is
 * withheld from what the viewer draws until it allows linked content, whatever
 * kind of content links it. A new kind of linked content the parser learns
 * fails here until the gate withholds it too.
 */
import JSZip from 'jszip';
import {
	ImageBuilder,
	MediaBuilder,
	PptxHandler,
	ShapeBuilder,
	TableBuilder,
} from 'ooxml-core/pptx';
import type { Model3DPptxElement, PptxData } from 'ooxml-core/pptx';
import { describe, expect, it } from 'vitest';

import { hasLinkedContent, isWebAddress, withholdLinkedContent } from './linked-content';

const ORIGIN = 'https://linked-content.test';
const CLICKED = 'https://clicked.test/page';
const PNG =
	'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAIAAAD91JpzAAAAEklEQVR4nGNkaGBgYGBgYGAAAAN4AE0r1L6tAAAAAElFTkSuQmCC';
const MP4 = 'data:video/mp4;base64,AAAAIGZ0eXBpc29tAAACAGlzb21pc28yYXZjMW1wNDE=';
const MP3 = 'data:audio/mpeg;base64,SUQzBAAAAAAAI1RTU0UAAAAPAAADTGF2ZjU4Ljc2LjEwMAAAAAAAAAAAAAAA';
const WAV = 'data:audio/wav;base64,UklGRiQAAABXQVZFZm10IBAAAAABAAEAQB8AAEAfAAABAAgAZGF0YQAAAAA=';
const GLB = 'data:model/gltf-binary;base64,Z2xURgIAAAAMAAAA';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const rel = (id: string, type: string, target: string) =>
	`<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}" TargetMode="External"/>`;

const sp = (id: number, name: string, inner: string, nvExtra = '') =>
	`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}">${nvExtra}</p:cNvPr><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${inner}`;

const FILL = `${sp(901, 'Linked fill', '<a:blipFill><a:blip r:link="rIdFill"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr></p:sp>')}`;
const BULLET = `${sp(902, 'Linked bullet', '</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr marL="342900" indent="-342900"><a:buBlip><a:blip r:embed="rIdBullet"/></a:buBlip></a:pPr><a:r><a:rPr lang="en-US"><a:hlinkClick r:id="rIdTextLink"/></a:rPr><a:t>https://written.test/in-text</a:t></a:r></a:p></p:txBody></p:sp>')}`;
const ACTIONS = `${sp(903, 'Action sounds', '</p:spPr></p:sp>', '<a:hlinkClick r:id="rIdShapeLink"><a:snd r:embed="rIdClickSnd" name="click.wav"/></a:hlinkClick><a:hlinkHover r:id="" action="ppaction://noaction"><a:snd r:embed="rIdHoverSnd" name="hover.wav"/></a:hlinkHover>')}`;
const BACKGROUND = `<p:bg><p:bgPr><a:blipFill><a:blip r:embed="rIdBackground"/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:effectLst/></p:bgPr></p:bg>`;
const CELL_FILL = `<a:blipFill><a:blip r:embed="rIdCell"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`;

/** A deck that links every kind of content it can, each from {@link ORIGIN}, and two hyperlinks. */
async function linkedDeck(): Promise<Uint8Array> {
	const { handler, data } = await PptxHandler.create({ initialSlideCount: 2 });
	const [first, second] = data.slides;
	first!.elements = [
		ImageBuilder.create(PNG).position(60, 80).size(300, 220).build(),
		MediaBuilder.video(MP4).position(420, 80).size(320, 180).build(),
		{
			id: 'linked-model',
			type: 'model3d',
			x: 780,
			y: 80,
			width: 160,
			height: 120,
			modelData: GLB,
			modelMimeType: 'model/gltf-binary',
			posterImage: PNG,
		} satisfies Model3DPptxElement,
	];
	const shape = ShapeBuilder.create('rect')
		.position(80, 80)
		.size(200, 100)
		.text('Animated')
		.build();
	second!.elements = [
		shape,
		TableBuilder.create().position(320, 80).size(400, 120).addRow(['A', 'B']).build(),
		MediaBuilder.audio(MP3).position(760, 80).size(60, 60).build(),
	];
	second!.transition = { type: 'fade', durationMs: 300, soundData: WAV } as never;
	second!.animations = [
		{ id: 'a1', elementId: shape.id, entrance: 'appear', trigger: 'onClick', soundData: WAV },
	] as never;
	for (const slide of data.slides) {
		slide.isDirty = true;
	}
	const zip = await JSZip.loadAsync(await handler.save(data.slides));
	handler.dispose();

	const edit = async (path: string, change: (xml: string) => string) => {
		zip.file(path, change(await zip.file(path)!.async('string')));
	};
	await edit('ppt/slides/slide1.xml', (xml) =>
		xml.replace('</p:spTree>', `${FILL}${BULLET}</p:spTree>`),
	);
	await edit('ppt/slides/slide2.xml', (xml) => {
		const withShapes = xml
			.replace(/<p:cSld([^>]*)>/u, `<p:cSld$1>${BACKGROUND}`)
			.replace('</p:spTree>', `${ACTIONS}</p:spTree>`);
		const cellFilled = withShapes.replace(/<a:tcPr([^>]*)\/>/u, `<a:tcPr$1>${CELL_FILL}</a:tcPr>`);
		return cellFilled === withShapes
			? withShapes.replace(/<a:tcPr([^>]*)>/u, `<a:tcPr$1>${CELL_FILL}`)
			: cellFilled;
	});
	const extra: Record<string, string> = {
		'ppt/slides/_rels/slide1.xml.rels':
			rel('rIdFill', 'image', `${ORIGIN}/fill.png`) +
			rel('rIdBullet', 'image', `${ORIGIN}/bullet.png`) +
			rel('rIdTextLink', 'hyperlink', CLICKED),
		'ppt/slides/_rels/slide2.xml.rels':
			rel('rIdBackground', 'image', `${ORIGIN}/background.png`) +
			rel('rIdCell', 'image', `${ORIGIN}/cell.png`) +
			rel('rIdClickSnd', 'audio', `${ORIGIN}/click.wav`) +
			rel('rIdHoverSnd', 'audio', `${ORIGIN}/hover.wav`) +
			rel('rIdShapeLink', 'hyperlink', CLICKED),
	};
	for (const [path, rels] of Object.entries(extra)) {
		// Link what the writer embedded (picture, video, poster, model poster, audio, the
		// transition and effect sounds) instead, and add the hand-written relationships.
		await edit(path, (xml) =>
			xml
				.replace(
					/(<Relationship\b[^>]*?)Target="\.\.\/media\/([^"]+\.(?:png|mp4|mp3|wav))"/gu,
					(_match, before: string, file: string) =>
						`${before}Target="${ORIGIN}/${file}" TargetMode="External"`,
				)
				.replace('</Relationships>', `${rels}</Relationships>`),
		);
	}
	for (const name of Object.keys(zip.files)) {
		if (/^ppt\/media\/.+\.(?:png|mp4|mp3|wav)$/u.test(name)) {
			zip.remove(name);
		}
	}
	return zip.generateAsync({ type: 'uint8array' });
}

async function loadLinkedDeck(): Promise<PptxData> {
	const handler = new PptxHandler();
	try {
		// What every viewer does: the model keeps the linked addresses.
		return await handler.load(await linkedDeck(), { allowExternalImages: true });
	} finally {
		handler.dispose();
	}
}

/** Every web address in `value`, with its model field, outside XML kept verbatim. */
function webAddresses(value: unknown): { field: string; url: string }[] {
	const found: { field: string; url: string }[] = [];
	const seen = new Set<object>();
	const walk = (item: unknown, key: string) => {
		if (isWebAddress(item)) {
			found.push({ field: key, url: item });
		} else if (typeof item === 'object' && item !== null && !seen.has(item)) {
			seen.add(item);
			for (const [childKey, child] of Object.entries(item)) {
				if (!childKey.startsWith('@_')) {
					walk(child, Array.isArray(item) ? key : childKey);
				}
			}
		}
	};
	walk(value, '');
	return found;
}

const isLinked = ({ url }: { url: string }) => URL.parse(url)?.origin === ORIGIN;

describe('linked content in a loaded deck', () => {
	it('links every kind of content the deck holds', async () => {
		const linked = webAddresses(await loadLinkedDeck()).filter(isLinked);
		const fields = new Set(linked.map(({ field }) => field));
		// Pictures, video, posters, fill, bullet, background, table cell, and every sound kind.
		expect([...fields].sort()).toStrictEqual(
			[
				'backgroundImage',
				'backgroundImageFillPath',
				'imageData',
				'imageDataUrl',
				'imagePath',
				'mediaPath',
				'posterImage',
				'soundPath',
			].sort(),
		);
		for (const sound of ['click.wav', 'hover.wav']) {
			expect(linked).toContainEqual({ field: 'soundPath', url: `${ORIGIN}/${sound}` });
		}
		// The transition and the effect sound, each linked from its own relationship.
		expect(
			linked.filter(({ field, url }) => field === 'soundPath' && /audio\d\.wav$/u.test(url)).length,
		).toBeGreaterThanOrEqual(2);
	});

	it('withholds every linked address from what a viewer draws before it allows them', async () => {
		const model = await loadLinkedDeck();
		expect(hasLinkedContent(model)).toBe(true);
		const view = withholdLinkedContent(model);
		expect(webAddresses(view).filter(isLinked)).toStrictEqual([]);
		expect(hasLinkedContent(view)).toBe(false);
	});

	it('keeps hyperlinks and text that only show or open an address', async () => {
		const model = await loadLinkedDeck();
		const kept = webAddresses(withholdLinkedContent(model)).filter((entry) => !isLinked(entry));
		expect(kept).toContainEqual({ field: 'url', url: CLICKED });
		expect(kept).toContainEqual({ field: 'hyperlink', url: CLICKED });
		expect(kept).toContainEqual({ field: 'text', url: 'https://written.test/in-text' });
	});
});
