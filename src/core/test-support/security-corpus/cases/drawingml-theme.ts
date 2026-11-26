import JSZip from 'jszip';
import { readThemePackage } from '../../../drawingml/theme-package';
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
];
