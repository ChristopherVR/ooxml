import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { NS } from '../xml/index';
import { themeMatrix } from './theme-matrix';
import { THEME_COLOR_SLOTS, type DrawingTheme } from './theme-model';
import { MAX_THEME_PACKAGE_BYTES, readThemePackage, writeThemePackage } from './theme-package';
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
