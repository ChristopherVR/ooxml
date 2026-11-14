import JSZip from 'jszip';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { PresentationBuilder } from '../../core/builders/sdk/PresentationBuilder';
import { PptxHandler } from '../../core/PptxHandler';
import type { PptxEmbeddedFont } from '../../core/types';
import { obfuscateFont } from '../../core/utils/font-deobfuscation';

/**
 * Embedded fonts honour the OS/2 `fsType` embedding permissions: a
 * restricted-licence face is never activated but its part is saved back
 * byte for byte; restricted, preview & print and bitmap-only faces are never
 * newly embedded.
 */
const FONT_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font';
const GUIDS = {
	restricted: '11111111-2222-3333-4444-555555555555',
	preview: '66666666-7777-8888-9999-AAAAAAAAAAAA',
	open: 'BBBBBBBB-CCCC-DDDD-EEEE-FFFFFFFFFFFF',
};

/** A minimal TrueType file whose OS/2 table carries `fsType`. */
function sfnt(fsType: number): Uint8Array {
	const bytes = new Uint8Array(12 + 2 * 16 + 96);
	const view = new DataView(bytes.buffer);
	view.setUint32(0, 0x00010000);
	view.setUint16(4, 2);
	view.setUint32(12, 0x68656164); // 'head'
	view.setUint32(28, 0x4f532f32); // 'OS/2'
	view.setUint32(28 + 8, 44);
	view.setUint32(28 + 12, 96);
	view.setUint16(44 + 8, fsType);
	return bytes;
}

const PARTS = {
	restricted: obfuscateFont(sfnt(0x0002), GUIDS.restricted),
	preview: obfuscateFont(sfnt(0x0004), GUIDS.preview),
	open: obfuscateFont(sfnt(0x0000), GUIDS.open),
};

async function blankDeck(): Promise<JSZip> {
	const { handler, data, createSlide } = await PresentationBuilder.create();
	data.slides.push(createSlide('Blank').build());
	return JSZip.loadAsync(await handler.save(data.slides));
}

async function deckWithFonts(): Promise<ArrayBuffer> {
	const zip = await blankDeck();
	const faces = (['restricted', 'preview', 'open'] as const).map(
		(kind, index) =>
			`<p:embeddedFont><p:font typeface="${kind} face"/><p:regular r:id="rIdF${index}" fontKey="{${GUIDS[kind]}}"/></p:embeddedFont>`,
	);
	const presentation = await zip.file('ppt/presentation.xml')!.async('string');
	zip.file(
		'ppt/presentation.xml',
		presentation.replace(
			'<p:defaultTextStyle>',
			`<p:embeddedFontLst>${faces.join('')}</p:embeddedFontLst><p:defaultTextStyle>`,
		),
	);
	const rels = await zip.file('ppt/_rels/presentation.xml.rels')!.async('string');
	const fontRels = (['restricted', 'preview', 'open'] as const).map(
		(kind, index) =>
			`<Relationship Id="rIdF${index}" Type="${FONT_REL_TYPE}" Target="fonts/font${index}.fntdata"/>`,
	);
	zip.file(
		'ppt/_rels/presentation.xml.rels',
		rels.replace('</Relationships>', `${fontRels.join('')}</Relationships>`),
	);
	zip.file('ppt/fonts/font0.fntdata', PARTS.restricted);
	zip.file('ppt/fonts/font1.fntdata', PARTS.preview);
	zip.file('ppt/fonts/font2.fntdata', PARTS.open);
	return (await zip.generateAsync({ type: 'uint8array' })).buffer as ArrayBuffer;
}

function byName(fonts: PptxEmbeddedFont[] | undefined, name: string): PptxEmbeddedFont {
	const font = fonts?.find((candidate) => candidate.name === name);
	expect(font, name).toBeDefined();
	return font!;
}

afterEach(() => {
	vi.restoreAllMocks();
});

describe('embedded font permissions', () => {
	it.each([false, true])(
		'does not activate a restricted face and warns (eager: %s)',
		async (eager) => {
			vi.spyOn(console, 'warn').mockImplementation(() => undefined);
			const handler = new PptxHandler();
			const data = await handler.load(await deckWithFonts(), {
				eagerDecodeEmbeddedFonts: eager,
			});
			const restricted = byName(data.embeddedFonts, 'restricted face');
			expect(restricted.fsType).toBe(0x0002);
			expect(restricted.dataUrl).toBe('');
			expect(restricted.rawFontData).toBeUndefined();
			await expect(restricted.load!()).resolves.toBeNull();

			const preview = byName(data.embeddedFonts, 'preview face');
			expect(preview.fsType).toBe(0x0004);
			expect(preview.rawFontData).toStrictEqual(sfnt(0x0004));
			expect(byName(data.embeddedFonts, 'open face').dataUrl).toMatch(/^data:font\/ttf;base64,/u);

			const warnings = (data.warnings ?? []).filter((w) => w.code === 'EMBEDDED_FONT_RESTRICTED');
			expect(warnings).toHaveLength(1);
			expect(warnings[0]!.message).toContain('restricted face');
			handler.dispose();
		},
	);

	it('saves the parts of restricted and preview faces back byte for byte', async () => {
		vi.spyOn(console, 'warn').mockImplementation(() => undefined);
		for (const mode of ['default', 'explicit', 'copied'] as const) {
			const handler = new PptxHandler();
			const data = await handler.load(await deckWithFonts());
			// Reading the preview face's data decodes it; it must still round-trip.
			void byName(data.embeddedFonts, 'preview face').rawFontData;
			const options =
				mode === 'default'
					? undefined
					: {
							embeddedFonts:
								mode === 'explicit'
									? data.embeddedFonts
									: // A host copy drops the non-enumerable runtime fields.
										data.embeddedFonts!.map((font) => ({ ...font })),
						};
			const saved = await JSZip.loadAsync(await handler.save(data.slides, options));
			expect(await saved.file('ppt/fonts/font0.fntdata')!.async('uint8array')).toStrictEqual(
				PARTS.restricted,
			);
			expect(await saved.file('ppt/fonts/font1.fntdata')!.async('uint8array')).toStrictEqual(
				PARTS.preview,
			);
			const fntdata = Object.keys(saved.files).filter((name) => name.endsWith('.fntdata'));
			expect(fntdata.sort()).toStrictEqual([
				'ppt/fonts/font0.fntdata',
				'ppt/fonts/font1.fntdata',
				'ppt/fonts/font2.fntdata',
			]);
			const presentation = await saved.file('ppt/presentation.xml')!.async('string');
			expect(presentation).toContain('typeface="restricted face"');
			expect(presentation).toContain('r:id="rIdF0"');
			const rels = await saved.file('ppt/_rels/presentation.xml.rels')!.async('string');
			expect(rels).toContain('Target="fonts/font0.fntdata"');
			handler.dispose();
		}
	});

	it('never newly embeds restricted, preview & print or bitmap-only faces', async () => {
		vi.spyOn(console, 'warn').mockImplementation(() => undefined);
		const zip = await blankDeck();
		const handler = new PptxHandler();
		const data = await handler.load(await zip.generateAsync({ type: 'arraybuffer' }));
		const host = (name: string, fsType: number): PptxEmbeddedFont => ({
			name,
			dataUrl: '',
			rawFontData: sfnt(fsType),
		});
		const saved = await JSZip.loadAsync(
			await handler.save(data.slides, {
				embeddedFonts: [
					host('New Restricted', 0x0002),
					host('New Preview', 0x0004),
					host('New Bitmap', 0x0200),
					host('New Editable', 0x0008),
				],
			}),
		);
		const fntdata = Object.keys(saved.files).filter((name) => name.endsWith('.fntdata'));
		expect(fntdata).toHaveLength(1);
		const presentation = await saved.file('ppt/presentation.xml')!.async('string');
		expect(presentation).toContain('typeface="New Editable"');
		expect(presentation).not.toContain('New Restricted');
		expect(presentation).not.toContain('New Preview');
		expect(presentation).not.toContain('New Bitmap');
		const skipped = handler
			.getCompatibilityWarnings()
			.filter((w) => w.code === 'EMBEDDED_FONT_NOT_EMBEDDABLE');
		expect(skipped.map((w) => w.message.match(/"(.+?)"/u)![1])).toStrictEqual([
			'New Restricted',
			'New Preview',
			'New Bitmap',
		]);
		handler.dispose();
	});

	it('writes no embedded font list when every new face is refused', async () => {
		vi.spyOn(console, 'warn').mockImplementation(() => undefined);
		const zip = await blankDeck();
		const handler = new PptxHandler();
		const data = await handler.load(await zip.generateAsync({ type: 'arraybuffer' }));
		const saved = await JSZip.loadAsync(
			await handler.save(data.slides, {
				embeddedFonts: [{ name: 'Only Restricted', dataUrl: '', rawFontData: sfnt(0x0002) }],
			}),
		);
		const presentation = await saved.file('ppt/presentation.xml')!.async('string');
		expect(presentation).not.toContain('embeddedFontLst');
		expect(Object.keys(saved.files).some((name) => name.endsWith('.fntdata'))).toBe(false);
		handler.dispose();
	});
});
