import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { themeFillMediaOf, themeSchemeMedia } from '../drawingml/theme-fill-media';
import { themeMatrix } from '../drawingml/theme-matrix';
import { THEME_COLOR_SLOTS, type DrawingTheme } from '../drawingml/theme-model';
import { readThemePackage, writeThemePackage } from '../drawingml/theme-package';
import { drawingThemeXml } from '../drawingml/write-theme';
import { NS } from '../xml/index';
import { createDocument } from './model';
import { loadDocx } from './parse';
import { saveDocx } from './save';
import { themeCatalogOf } from './theme';

const png = new Uint8Array([137, 80, 78, 71, 13, 10, 26, 10, 4, 5, 6]);
const theme: DrawingTheme = {
	name: 'Textured',
	colorScheme: {
		colors: Object.fromEntries(
			THEME_COLOR_SLOTS.map((slot) => [slot, { kind: 'srgb', value: '336699', transforms: [] }]),
		),
	},
	fontScheme: {
		major: { latin: 'Georgia', scripts: {} },
		minor: { latin: 'Verdana', scripts: {} },
	},
	formatScheme: themeMatrix(),
};

/** A `.thmx` whose first fill style is a tiled picture fill. */
async function texturedThemePackage(): Promise<Uint8Array> {
	const xml = drawingThemeXml(theme).replace(
		/<a:fillStyleLst><a:solidFill>.*?<\/a:solidFill>/u,
		`<a:fillStyleLst><a:blipFill><a:blip xmlns:r="${NS.r}" r:embed="rId9"/><a:tile/></a:blipFill>`,
	);
	return writeThemePackage(xml, { media: new Map([['rId9', themeFillMediaOf(png, 'image/png')]]) });
}

describe('Word theme picture fills', () => {
	it('saves an imported theme picture fill with its picture and reads it back', async () => {
		const imported = await readThemePackage(await texturedThemePackage());
		const loaded = await loadDocx(await saveDocx(createDocument()));
		const next = structuredClone(loaded.model);
		next.theme = { ...themeCatalogOf(imported), colorMapping: next.theme?.colorMapping ?? {} };
		const bytes = await loaded.save(next);
		const zip = await JSZip.loadAsync(bytes);
		const rels = await zip.file('word/theme/_rels/theme1.xml.rels')?.async('string');
		const id = /Id="([^"]+)"[^>]*Target="\.\.\/media\/(image\d+\.png)"/u.exec(rels ?? '');
		expect(id).not.toBeNull();
		expect(await zip.file(`word/media/${id![2]}`)?.async('uint8array')).toEqual(png);
		expect(await zip.file('word/theme/theme1.xml')?.async('string')).toContain(
			`r:embed="${id![1]}"`,
		);
		const reopened = await loadDocx(bytes);
		const fill = reopened.model.theme?.formatScheme?.fillStyles[0];
		expect(fill?.fill.kind).toBe('picture');
		expect(themeSchemeMedia(reopened.model.theme?.formatScheme).get(id![1]!)?.contentType).toBe(
			'image/png',
		);
	});

	it('keeps the picture relationship when only the theme colours change', async () => {
		const imported = await readThemePackage(await texturedThemePackage());
		const first = await loadDocx(await saveDocx(createDocument()));
		const withTheme = structuredClone(first.model);
		withTheme.theme = { ...themeCatalogOf(imported), colorMapping: {} };
		const reopened = await loadDocx(await first.save(withTheme));
		const recoloured = structuredClone(reopened.model);
		recoloured.theme = {
			...recoloured.theme!,
			colors: { ...recoloured.theme!.colors, accent1: 'FF0000' },
		};
		const zip = await JSZip.loadAsync(await reopened.save(recoloured));
		expect(
			Object.keys(zip.files).filter((name) => /^word\/media\/image\d+\.png$/u.test(name)),
		).toHaveLength(1);
		expect(await zip.file('word/theme/theme1.xml')?.async('string')).toContain('a:blipFill');
	});
});
