import JSZip from 'jszip';
import { readThemePackage, writeThemePackage } from '../../../drawingml/theme-package';
import { themeFillMediaOf } from '../../../drawingml/theme-fill-media';
import { themeMatrix } from '../../../drawingml/theme-matrix';
import { THEME_COLOR_SLOTS } from '../../../drawingml/theme-model';
import { drawingThemeXml } from '../../../drawingml/write-theme';
import { ZipLimitError } from '../../../opc/zip-limits';
import { NS } from '../../../xml/index';
import type { SecurityCase } from '../case';

export const cases: readonly SecurityCase[] = [
	{
		name: 'standalone theme XML is limited by bytes before it is parsed',
		limits: { ms: 1_000, heapMb: 64 },
		async run() {
			const zip = new JSZip();
			zip.file(
				'[Content_Types].xml',
				`<Types xmlns="${NS.contentTypes}"/><!--${'界'.repeat(175_000)}-->`,
			);
			try {
				await readThemePackage(
					await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }),
				);
				return 'accepted';
			} catch (error) {
				return error instanceof Error ? error.message : String(error);
			}
		},
		expected: 'Office theme: invalid XML part size.',
	},
	{
		name: 'standalone theme enforces the XML stream limit when size metadata disagrees',
		limits: { ms: 1_000, heapMb: 64 },
		async run() {
			const name = '[Content_Types].xml';
			const zip = new JSZip();
			zip.file(name, `<Types xmlns="${NS.contentTypes}"/><!--${' '.repeat(600_000)}-->`);
			const input = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
			// Simulate inconsistent metadata at the ZIP decoder boundary using an ordinary archive.
			const load = JSZip.loadAsync;
			JSZip.loadAsync = async (...args: Parameters<typeof load>) => {
				const loaded = await load(...args);
				const entry = loaded.file(name);
				const metadata = entry && '_data' in entry ? entry._data : undefined;
				if (!metadata || typeof metadata !== 'object' || !('uncompressedSize' in metadata))
					throw new Error('Missing decoded size metadata');
				metadata.uncompressedSize = 128;
				return loaded;
			};
			try {
				await readThemePackage(input);
				return 'accepted';
			} catch (error) {
				return error instanceof ZipLimitError
					? { kind: error.kind, limit: error.limit }
					: String(error);
			} finally {
				JSZip.loadAsync = load;
			}
		},
		expected: { kind: 'declaredSize', limit: 512 * 1024 },
	},
	{
		name: 'theme picture fills never inflate a picture past the per-picture limit',
		limits: { ms: 3_000, heapMb: 96 },
		async run() {
			const xml = drawingThemeXml({
				name: 'Pictures',
				colorScheme: {
					colors: Object.fromEntries(
						THEME_COLOR_SLOTS.map((slot) => [
							slot,
							{ kind: 'srgb', value: '000000', transforms: [] },
						]),
					),
				},
				fontScheme: { major: { latin: 'A', scripts: {} }, minor: { latin: 'B', scripts: {} } },
				formatScheme: themeMatrix(),
			}).replace(
				/<a:fillStyleLst><a:solidFill>.*?<\/a:solidFill>/u,
				`<a:fillStyleLst><a:blipFill><a:blip xmlns:r="${NS.r}" r:embed="rId1"/></a:blipFill>`,
			);
			const media = new Map([
				['rId1', themeFillMediaOf(new Uint8Array([137, 80, 78, 71]), 'image/png')],
			]);
			const zip = await JSZip.loadAsync(await writeThemePackage(xml, { media }));
			zip.file(
				'theme/theme/_rels/theme1.xml.rels',
				`<Relationships xmlns="${NS.rels}"><Relationship Id="rId1" Type="${NS.r}/image" Target="../media/big.png"/></Relationships>`,
			);
			// Highly compressible: 6 MB of zeros deflate to a few kilobytes.
			zip.file('theme/media/big.png', new Uint8Array(6 * 1024 * 1024));
			const warnings: string[] = [];
			const theme = await readThemePackage(
				await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }),
				{ onWarning: (warning) => warnings.push(warning) },
			);
			return { warnings, fill: theme.formatScheme?.fillStyles[0]?.fill.kind };
		},
		expected: { warnings: ['pictureFillsReplaced'], fill: 'solid' },
	},
];
