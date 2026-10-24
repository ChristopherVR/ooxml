/**
 * Integration: the background a slide master and its layouts load with, and the background a
 * layout preview paints, resolve through the master's own colour map, the way a slide's does.
 *
 * A master's `p:bg` is usually `<p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef>`, and
 * what `bg1` means is decided by the master's `p:clrMap`: a dark background style (Design >
 * Background Styles, the dark columns) writes `bg1="dk1"` there. The master and layout parsers
 * resolved the reference with no colour map active (the default `bg1 -> lt1`), so a dark master
 * loaded white while its own slides loaded black, and a layout preview resolved its background
 * and its shapes through whichever master happened to be active last.
 */
import fs from 'node:fs';
import path from 'node:path';

import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';

import { duplicateSlideMaster } from '../../core/builders/sdk';
import { PptxHandler } from '../../core/PptxHandler';
import type { PptxData, PptxElement } from '../../core/types';
import { requireFixture } from '../require-fixture';

/** One master (`bgRef idx="1001"` on `bg1`, light colour map) and eleven layouts. */
const FIXTURE = requireFixture(path.resolve(__dirname, '../fixtures/theme-styles.pptx'));

const MASTER = 'ppt/slideMasters/slideMaster1.xml';
const TITLE_LAYOUT = 'ppt/slideLayouts/slideLayout1.xml';
const BLANK_LAYOUT = 'ppt/slideLayouts/slideLayout7.xml';
const LIGHT_MAP = 'bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2"';
const DARK_MAP = 'bg1="dk1" tx1="lt1" bg2="dk2" tx2="lt2"';
const ACCENTS =
	'accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" ' +
	'accent6="accent6" hlink="hlink" folHlink="folHlink"';
const OWN_BACKGROUND = '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>';
/** A plain rectangle filled with `tx1`, so a preview shows which colour map drew it. */
const TX1_RECT =
	'<p:sp><p:nvSpPr><p:cNvPr id="90" name="Band"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>' +
	'<a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>' +
	'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="tx1"/>' +
	'</a:solidFill></p:spPr></p:sp>';

function fixtureBytes(): ArrayBuffer {
	const file = fs.readFileSync(FIXTURE);
	return file.buffer.slice(file.byteOffset, file.byteOffset + file.byteLength) as ArrayBuffer;
}

function arrayBuffer(bytes: Uint8Array): ArrayBuffer {
	return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

/** Apply `edits` (part path to a rewrite of its XML) to a package and return the new bytes. */
async function rewrite(
	source: ArrayBuffer,
	edits: Record<string, (xml: string) => string>,
): Promise<ArrayBuffer> {
	const zip = await JSZip.loadAsync(source);
	for (const [part, edit] of Object.entries(edits)) {
		const xml = await zip.file(part)!.async('string');
		const next = edit(xml);
		expect(next, `the edit of ${part} changed nothing`).not.toBe(xml);
		zip.file(part, next);
	}
	return arrayBuffer(await zip.generateAsync({ type: 'uint8array' }));
}

const darkMap = (xml: string): string => xml.replace(LIGHT_MAP, DARK_MAP);
const withOwnBackground = (xml: string): string =>
	xml.replace(/(<p:cSld name="[^"]*">)/, `$1${OWN_BACKGROUND}`);
const withTx1Rect = (xml: string): string => xml.replace('</p:grpSpPr>', `</p:grpSpPr>${TX1_RECT}`);

async function load(bytes: ArrayBuffer): Promise<{ handler: PptxHandler; data: PptxData }> {
	const handler = new PptxHandler();
	return { handler, data: await handler.load(bytes) };
}

async function previewColour(handler: PptxHandler, layoutPath: string): Promise<string> {
	const preview = await handler.getLayoutPreview(layoutPath);
	expect(preview, `a preview of ${layoutPath}`).not.toBeNull();
	return preview!.backgroundColor?.toUpperCase() ?? '';
}

function rectFill(elements: readonly PptxElement[]): string | undefined {
	const rect = elements.find((element) => element.name === 'Band');
	expect(rect, 'the injected tx1 rectangle').toBeDefined();
	return 'shapeStyle' in rect! ? rect.shapeStyle?.fillColor?.toUpperCase() : undefined;
}

describe('master and layout backgrounds at load', () => {
	it('loads a master saved with a dark background style dark, like its slides', async () => {
		const { handler, data } = await load(fixtureBytes());
		const dark = handler.applyBackgroundStyleToSlides(data.slides, {}, 4, { kind: 'all' })!;
		await handler.commitMasterBackgroundStyle(4);
		const saved = await handler.save(dark.slides);

		const reopened = await new PptxHandler().load(arrayBuffer(saved));
		expect(reopened.slides[0]!.backgroundColor?.toUpperCase()).toBe('#000000');
		expect(reopened.slideMasters?.[0]?.backgroundColor?.toUpperCase()).toBe('#000000');
	});

	it("resolves a layout's own background reference through its master's colour map", async () => {
		const { data } = await load(
			await rewrite(fixtureBytes(), { [MASTER]: darkMap, [BLANK_LAYOUT]: withOwnBackground }),
		);
		const master = data.slideMasters![0]!;
		expect(master.backgroundColor?.toUpperCase()).toBe('#000000');
		const blank = master.layouts?.find((layout) => layout.path === BLANK_LAYOUT);
		expect(blank?.backgroundColor?.toUpperCase()).toBe('#000000');
		// A layout without its own `p:bg` still reports none: it inherits the master's.
		const title = master.layouts?.find((layout) => layout.path === TITLE_LAYOUT);
		expect(title?.backgroundColor).toBeUndefined();
	});

	it('keeps a light master light', async () => {
		const { data } = await load(fixtureBytes());
		expect(data.slideMasters?.[0]?.backgroundColor?.toUpperCase()).toBe('#FFFFFF');
	});
});

describe('layout previews', () => {
	it("paint a dark master's layouts dark, inherited or their own", async () => {
		const { handler } = await load(
			await rewrite(fixtureBytes(), { [MASTER]: darkMap, [BLANK_LAYOUT]: withOwnBackground }),
		);
		expect(await previewColour(handler, TITLE_LAYOUT)).toBe('#000000');
		expect(await previewColour(handler, BLANK_LAYOUT)).toBe('#000000');
	});

	it("follow a layout's own colour map override, as its shapes do", async () => {
		const lightOverride = `<p:clrMapOvr><a:overrideClrMapping ${LIGHT_MAP} ${ACCENTS}/></p:clrMapOvr>`;
		const { handler } = await load(
			await rewrite(fixtureBytes(), {
				[MASTER]: darkMap,
				[TITLE_LAYOUT]: (xml) =>
					withTx1Rect(xml).replace(
						'<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>',
						lightOverride,
					),
			}),
		);
		expect(await previewColour(handler, TITLE_LAYOUT)).toBe('#FFFFFF');
		expect(rectFill((await handler.getLayoutPreview(TITLE_LAYOUT))!.elements)).toBe('#000000');
		expect(await previewColour(handler, BLANK_LAYOUT)).toBe('#000000');
	});
});

describe('a deck with a light and a dark master', () => {
	async function twoMasters(): Promise<{ handler: PptxHandler; data: PptxData }> {
		const seed = await load(fixtureBytes());
		const duplicated = await duplicateSlideMaster(seed.handler, seed.data, MASTER);
		expect(duplicated.ok).toBe(true);
		if (!duplicated.ok) {
			throw new Error('duplicateSlideMaster failed');
		}
		const bytes = arrayBuffer(await duplicated.handler.save(duplicated.data.slides));
		const second = duplicated.masterId;
		const secondLayout = duplicated.data.slideMasters!.find((m) => m.path === second)!
			.layoutPaths![0]!;
		return load(
			await rewrite(bytes, {
				[second]: darkMap,
				[secondLayout]: withTx1Rect,
				[TITLE_LAYOUT]: withTx1Rect,
			}),
		);
	}

	it('loads each master through its own colour map', async () => {
		const { data } = await twoMasters();
		const [light, dark] = data.slideMasters!;
		expect(light!.backgroundColor?.toUpperCase()).toBe('#FFFFFF');
		expect(dark!.backgroundColor?.toUpperCase()).toBe('#000000');
	});

	// Previews are memoised, so each order gets a handler of its own.
	it.each(['dark first', 'light first'])(
		'previews each layout through its own master (%s)',
		async (order) => {
			const { handler, data } = await twoMasters();
			const darkLayout = data.slideMasters![1]!.layoutPaths![0]!;
			const paths =
				order === 'dark first' ? [darkLayout, TITLE_LAYOUT] : [TITLE_LAYOUT, darkLayout];
			const previews = await handler.getLayoutPreviews(paths);
			for (const preview of previews) {
				const dark = preview.path === darkLayout;
				expect(preview.backgroundColor?.toUpperCase()).toBe(dark ? '#000000' : '#FFFFFF');
				// tx1 is light under the dark map and dark under the light one.
				expect(rectFill(preview.elements)).toBe(dark ? '#FFFFFF' : '#000000');
			}
		},
	);
});
