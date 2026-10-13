/**
 * `p:presentation/@firstSlideNum` (ECMA-376 §19.2.1.26) sets the number a
 * `slidenum` field shows on the first slide; PowerPoint numbers slide `i`
 * (0-based) as `firstSlideNum + i`. The loader exposes it as
 * `PptxData.firstSlideNumber` for the renderers, and a save must keep it.
 */
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { serializePptxToJson } from '../../converter/json';
import { PresentationBuilder } from '../../core/builders/sdk/PresentationBuilder';
import { PptxHandler } from '../../core/PptxHandler';

/** A one-slide deck whose `p:presentation` root carries `extraAttributes`. */
async function deckWithPresentationAttributes(extraAttributes: string): Promise<Uint8Array> {
	const built = await PresentationBuilder.create({ initialSlideCount: 1 });
	const seed = await built.handler.save(built.data.slides);
	const zip = await JSZip.loadAsync(seed);
	const presentationXml = await zip.file('ppt/presentation.xml')!.async('string');
	const patched = presentationXml.replace(
		'saveSubsetFonts="1">',
		`saveSubsetFonts="1" ${extraAttributes}>`,
	);
	expect(patched).not.toBe(presentationXml);
	zip.file('ppt/presentation.xml', patched);
	return zip.generateAsync({ type: 'uint8array' });
}

async function load(bytes: Uint8Array) {
	const handler = new PptxHandler();
	const data = await handler.load(bytes.buffer as ArrayBuffer);
	return { handler, data };
}

describe('p:presentation/@firstSlideNum', () => {
	it('is undefined (spec default: 1) when the attribute is absent', async () => {
		const { data } = await load(await deckWithPresentationAttributes(''));
		expect(data.firstSlideNumber).toBeUndefined();
	});

	it('parses an explicit @firstSlideNum onto the typed model', async () => {
		const { data } = await load(await deckWithPresentationAttributes('firstSlideNum="45"'));
		expect(data.firstSlideNumber).toBe(45);
		// The document position is unchanged: navigation still uses it.
		expect(data.slides[0]!.slideNumber).toBe(1);
	});

	it('accepts zero and ignores a value that is not an integer', async () => {
		const zero = await load(await deckWithPresentationAttributes('firstSlideNum="0"'));
		expect(zero.data.firstSlideNumber).toBe(0);
		const junk = await load(await deckWithPresentationAttributes('firstSlideNum="4.5"'));
		expect(junk.data.firstSlideNumber).toBeUndefined();
	});

	it('accepts the xsd:int bounds and rejects values outside them', async () => {
		const firstSlideNumber = async (lexical: string) =>
			(await load(await deckWithPresentationAttributes(`firstSlideNum="${lexical}"`))).data
				.firstSlideNumber;
		expect(await firstSlideNumber('2147483647')).toBe(2_147_483_647);
		expect(await firstSlideNumber('-2147483648')).toBe(-2_147_483_648);
		expect(await firstSlideNumber('2147483648')).toBeUndefined();
		expect(await firstSlideNumber('-2147483649')).toBeUndefined();
		// Parses to Infinity / -Infinity.
		expect(await firstSlideNumber('9'.repeat(400))).toBeUndefined();
		expect(await firstSlideNumber(`-${'9'.repeat(400)}`)).toBeUndefined();
	});

	it('survives a save that rewrites presentation.xml and reloads as the same value', async () => {
		const { handler, data } = await load(
			await deckWithPresentationAttributes('firstSlideNum="45"'),
		);
		// Adding a slide changes the sldIdLst, so presentation.xml is rewritten.
		const extra = await PresentationBuilder.create({ initialSlideCount: 1 });
		data.slides.push(extra.data.slides[0]!);
		const saved = await handler.save(data.slides);

		const zip = await JSZip.loadAsync(saved);
		const presentationXml = await zip.file('ppt/presentation.xml')!.async('string');
		expect(presentationXml).toContain('firstSlideNum="45"');

		const reloaded = await load(saved);
		expect(reloaded.data.firstSlideNumber).toBe(45);
		expect(reloaded.data.slides).toHaveLength(2);
	});

	it('keeps the number through JSON import -> save -> reopen', async () => {
		const { data } = await load(await deckWithPresentationAttributes('firstSlideNum="45"'));
		const json = new TextEncoder().encode(serializePptxToJson(data));
		const handler = new PptxHandler();
		const imported = await handler.load(json.buffer as ArrayBuffer);
		expect(imported.firstSlideNumber).toBe(45);

		const saved = await handler.save(imported.slides);
		const zip = await JSZip.loadAsync(saved);
		const presentationXml = await zip.file('ppt/presentation.xml')!.async('string');
		expect(presentationXml).toContain('firstSlideNum="45"');

		const reopened = await load(saved);
		expect(reopened.data.firstSlideNumber).toBe(45);
	});
});
