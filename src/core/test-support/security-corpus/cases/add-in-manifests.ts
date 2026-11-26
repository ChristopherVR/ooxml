import JSZip from 'jszip';
import { AddInManifestError, readLocalAddIn } from '../../../automation/add-ins';
import type { SecurityCase } from '../case';

async function refusal(bytes: Uint8Array): Promise<string> {
	try {
		await readLocalAddIn(bytes, { host: 'word' });
		return 'accepted';
	} catch (error) {
		if (error instanceof AddInManifestError) return error.code;
		throw error;
	}
}

export const cases: readonly SecurityCase[] = [
	{
		name: 'external entity declarations are rejected offline',
		limits: { ms: 2_000, heapMb: 64 },
		run: () =>
			refusal(
				new TextEncoder().encode(
					'<!DOCTYPE OfficeApp SYSTEM "https://example.invalid/manifest.dtd"><OfficeApp/>',
				),
			),
		expected: 'manifest',
	},
	{
		name: 'escaped duplicate keys cannot override a reviewed property',
		limits: { ms: 2_000, heapMb: 64 },
		run: () => refusal(new TextEncoder().encode('{"id":"a","\\u0069d":"b"}')),
		expected: 'manifest',
	},
	{
		name: 'the ZIP directory is inspected before duplicate paths can collapse',
		limits: { ms: 2_000, heapMb: 64 },
		async run() {
			const zip = new JSZip().file('manifest.json', '{}').file('manifest.jso0', '{}');
			const data = await zip.generateAsync({ type: 'nodebuffer' });
			const alternate = 'manifest.jso0';
			for (
				let offset = data.indexOf(alternate);
				offset >= 0;
				offset = data.indexOf(alternate, offset + 1)
			) {
				data.write('manifest.json', offset, 'utf8');
			}
			return refusal(data);
		},
		expected: 'package',
	},
	{
		name: 'archive paths cannot traverse through URI decoding',
		limits: { ms: 2_000, heapMb: 64 },
		async run() {
			const zip = new JSZip().file('%2e%2e/manifest.json', '{}', { createFolders: false });
			return refusal(await zip.generateAsync({ type: 'uint8array' }));
		},
		expected: 'package',
	},
];
