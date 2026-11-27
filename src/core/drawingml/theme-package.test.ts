import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { NS } from '../xml/index';
import { themeMatrix } from './theme-matrix';
import { THEME_COLOR_SLOTS, type DrawingTheme } from './theme-model';
import { MAX_THEME_PACKAGE_BYTES, readThemePackage, writeThemePackage } from './theme-package';
import { themeFillMediaOf, themeSchemeMedia } from './theme-fill-media';
import { drawingThemeXml } from './write-theme';

const theme: DrawingTheme = {
	name: 'Reading room',
	colorScheme: {
		name: 'Ink',
		colors: Object.fromEntries(
			THEME_COLOR_SLOTS.map((slot) => [
				slot,
				{ kind: 'srgb', value: slot === 'accent1' ? '123456' : '000000', transforms: [] },
			]),
		),
	},
	fontScheme: {
		name: 'Reading',
		major: { latin: 'Georgia', scripts: { Jpan: 'Yu Mincho' } },
		minor: { latin: 'Verdana', scripts: {} },
	},
	formatScheme: themeMatrix(),
};
const part = 'theme/theme/theme1.xml';
const managerRels = 'theme/theme/_rels/themeManager.xml.rels';
const bytes = () => writeThemePackage(drawingThemeXml(theme));
const withPictureFills = (xml: string) =>
	xml.replace(
		/<a:(fillStyleLst|bgFillStyleLst)><a:solidFill>.*?<\/a:solidFill>/gu,
		'<a:$1><a:blipFill><a:blip xmlns:media="' + NS.r + '" media:embed="rId2"/></a:blipFill>',
	);

async function replacePart(name: string, change: (xml: string) => string) {
	const zip = await JSZip.loadAsync(await bytes());
	const xml = await zip.file(name)?.async('string');
	if (!xml) throw new Error(`Missing test part ${name}`);
	zip.file(name, change(xml));
	return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

describe('standalone Office themes', () => {
	it('round trips the schemes through the theme manager relationship', async () => {
		const loaded = await readThemePackage(await bytes());
		expect(loaded).toMatchObject(theme);
		expect(loaded.formatScheme?.effectStyles).toHaveLength(3);
	});

	it('uses the related part when the DrawingML part has a different name', async () => {
		const zip = await JSZip.loadAsync(await bytes());
		const manifest = await zip.file('[Content_Types].xml')?.async('string');
		const rels = await zip.file(managerRels)?.async('string');
		if (!manifest || !rels) throw new Error('Missing theme registration');
		zip.file('[Content_Types].xml', manifest.replace(part, 'theme/theme/schemes.xml'));
		zip.file(managerRels, rels.replace('theme1.xml', 'schemes.xml'));
		zip.file('theme/theme/schemes.xml', drawingThemeXml(theme));
		zip.remove(part);
		expect((await readThemePackage(await zip.generateAsync({ type: 'uint8array' }))).name).toBe(
			theme.name,
		);
	});

	it('accepts a UTF-16 theme part', async () => {
		const zip = await JSZip.loadAsync(await bytes());
		const xml = '\uFEFF' + drawingThemeXml(theme).replace('UTF-8', 'UTF-16');
		const encoded = new Uint8Array(xml.length * 2);
		const view = new DataView(encoded.buffer);
		for (let i = 0; i < xml.length; i++) view.setUint16(i * 2, xml.charCodeAt(i), true);
		zip.file(part, encoded);
		expect(await readThemePackage(await zip.generateAsync({ type: 'uint8array' }))).toMatchObject(
			theme,
		);
	});

	it('writes UTF-8 bytes even when the source text declares another encoding', async () => {
		const xml = drawingThemeXml(theme).replace('UTF-8', 'UTF-16');
		const zip = await JSZip.loadAsync(await writeThemePackage(xml));
		expect(await zip.file(part)?.async('string')).not.toContain('UTF-16');
		expect(await readThemePackage(await writeThemePackage(xml))).toMatchObject(theme);
	});

	it('replaces picture fills in both lists on import and warns once', async () => {
		const warnings: string[] = [];
		const loaded = await readThemePackage(await replacePart(part, withPictureFills), {
			onWarning: (warning) => warnings.push(warning),
		});
		expect(warnings).toEqual(['pictureFillsReplaced']);
		expect(loaded).toMatchObject(theme);
		expect(loaded.formatScheme?.sourceXml).not.toContain('rId2');
		expect(loaded.formatScheme?.fillStyles).toHaveLength(3);
		expect(loaded.formatScheme?.backgroundFillStyles).toHaveLength(3);
	});

	it('exports a self-contained package with picture fills replaced and a warning', async () => {
		const warnings: string[] = [];
		const output = await writeThemePackage(withPictureFills(drawingThemeXml(theme)), {
			onWarning: (warning) => warnings.push(warning),
		});
		expect(
			await readThemePackage(output, { onWarning: (warning) => warnings.push(warning) }),
		).toMatchObject(theme);
		expect(warnings).toEqual(['pictureFillsReplaced']);
		const zip = await JSZip.loadAsync(output);
		expect(await zip.file(part)?.async('string')).not.toContain('blipFill');
	});

	it('admits Office packages with variant parts within the package part cap', async () => {
		const zip = await JSZip.loadAsync(await bytes());
		for (let index = 0; index < 568; index++) {
			zip.file(`themeVariants/part${index}.xml`, '<variant/>', { createFolders: false });
		}
		expect(await readThemePackage(await zip.generateAsync({ type: 'uint8array' }))).toMatchObject(
			theme,
		);
		for (let index = 568; index < 640; index++) {
			zip.file(`themeVariants/part${index}.xml`, '<variant/>', { createFolders: false });
		}
		await expect(
			readThemePackage(await zip.generateAsync({ type: 'uint8array' })),
		).rejects.toMatchObject({
			kind: 'maxEntries',
		});
	});

	it.each([
		['wrong root', part, () => `<a:themeOverride xmlns:a="${NS.a}"/>`, /expected theme/],
		['DTD', part, (xml: string) => xml.replace('<a:theme ', '<!DOCTYPE a:theme><a:theme '), /DTD/],
		[
			'missing colour slot',
			part,
			(xml: string) => xml.replaceAll('accent1', 'wrong'),
			/expected accent1/,
		],
		['invalid RGB', part, (xml: string) => xml.replace('123456', 'INVALID'), /incomplete schemes/],
		[
			'invalid fill style',
			part,
			(xml: string) => xml.replaceAll('a:solidFill', 'a:fakeFill'),
			/invalid style matrix/,
		],
		[
			'deep XML',
			part,
			(xml: string) =>
				xml.replace(
					'<a:objectDefaults/>',
					'<a:objectDefaults>'.repeat(41) + '</a:objectDefaults>'.repeat(41),
				),
			/nesting depth/,
		],
		[
			'related media',
			part,
			(xml: string) => xml.replace('<a:fmtScheme', `<a:fmtScheme xmlns:r="${NS.r}" r:id="rId1"`),
			/related resources/,
		],
		[
			'external theme',
			managerRels,
			(xml: string) =>
				xml.replace(
					'Target="theme1.xml"',
					'TargetMode="External" Target="https://example.invalid/theme1.xml"',
				),
			/internal relationship/,
		],
		[
			'escaping target',
			managerRels,
			(xml: string) => xml.replace('theme1.xml', '../../../theme1.xml'),
			/invalid part target/,
		],
	])('rejects %s', async (_label, name, change, reason) => {
		await expect(readThemePackage(await replacePart(name, change))).rejects.toThrow(reason);
	});

	it('rejects an oversized input, an inflated entry and an oversized XML part', async () => {
		await expect(
			readThemePackage(new Uint8Array(MAX_THEME_PACKAGE_BYTES + 1)),
		).rejects.toMatchObject({ kind: 'maxInputBytes' });
		await expect(
			readThemePackage(await replacePart(part, () => ' '.repeat(MAX_THEME_PACKAGE_BYTES + 1))),
		).rejects.toMatchObject({ kind: 'maxEntryBytes' });
		await expect(
			readThemePackage(await replacePart(part, () => ' '.repeat(512 * 1024 + 1))),
		).rejects.toThrow(/invalid XML part size/);
		await expect(
			readThemePackage(
				await replacePart(part, (xml) => xml + '<!--' + '界'.repeat(175_000) + '-->'),
			),
		).rejects.toThrow(/invalid XML part size/);
	});
});

describe('standalone Office theme picture fills', () => {
	const png = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10, 1, 2, 3]);
	const media = new Map([['rId2', themeFillMediaOf(png, 'image/png')]]);

	it('keeps a picture fill with its picture through export and import', async () => {
		const warnings: string[] = [];
		const output = await writeThemePackage(withPictureFills(drawingThemeXml(theme)), {
			media,
			onWarning: (warning) => warnings.push(warning),
		});
		const zip = await JSZip.loadAsync(output);
		expect(await zip.file('theme/media/image1.png')?.async('uint8array')).toEqual(png);
		expect(await zip.file('theme/theme/_rels/theme1.xml.rels')?.async('string')).toContain(
			'../media/image1.png',
		);
		const loaded = await readThemePackage(output, {
			onWarning: (warning) => warnings.push(warning),
		});
		expect(warnings).toEqual([]);
		expect(loaded.formatScheme?.fillStyles[0]?.fill.kind).toBe('picture');
		expect(loaded.formatScheme?.backgroundFillStyles[0]?.fill.kind).toBe('picture');
		expect(themeSchemeMedia(loaded.formatScheme).get('rId2')).toEqual(media.get('rId2'));
	});

	it('falls back to a solid fill for a picture of an unsupported type', async () => {
		const warnings: string[] = [];
		const output = await writeThemePackage(withPictureFills(drawingThemeXml(theme)), {
			media: new Map([['rId2', themeFillMediaOf(png, 'image/svg+xml')]]),
			onWarning: (warning) => warnings.push(warning),
		});
		expect(warnings).toEqual(['pictureFillsReplaced']);
		const zip = await JSZip.loadAsync(output);
		expect(await zip.file(part)?.async('string')).not.toContain('blipFill');
	});

	it('ignores picture relationships that leave the package', async () => {
		const output = await writeThemePackage(withPictureFills(drawingThemeXml(theme)), { media });
		const zip = await JSZip.loadAsync(output);
		zip.file(
			'theme/theme/_rels/theme1.xml.rels',
			`<Relationships xmlns="${NS.rels}"><Relationship Id="rId2" Type="${NS.r}/image" Target="https://example.invalid/a.png" TargetMode="External"/></Relationships>`,
		);
		const warnings: string[] = [];
		const loaded = await readThemePackage(await zip.generateAsync({ type: 'uint8array' }), {
			onWarning: (warning) => warnings.push(warning),
		});
		expect(warnings).toEqual(['pictureFillsReplaced']);
		expect(loaded.formatScheme?.fillStyles[0]?.fill.kind).toBe('solid');
	});
});
