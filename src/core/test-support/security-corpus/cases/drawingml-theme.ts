import JSZip from 'jszip';
import { readThemePackage } from '../../../drawingml/theme-package';
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
];
