import JSZip from 'jszip';
import { AddInManifestError, readLocalAddIn } from '../../../automation/add-ins';
import type { SecurityCase } from '../case';

async function refusal(bytes: Uint8Array, report: 'code' | 'field' = 'code') {
	try {
		await readLocalAddIn(bytes, { host: 'word' });
		return 'accepted';
	} catch (error) {
		if (error instanceof AddInManifestError) return error[report];
		throw error;
	}
}

const unsafePane = `<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
	xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
	<Id>083a367e-842b-4d4f-ab91-cd5de849e495</Id><Version>1.0.0</Version>
	<ProviderName>Fixture</ProviderName><DefaultLocale>en-US</DefaultLocale>
	<DisplayName DefaultValue="Example"/><Description DefaultValue="Example"/>
	<Hosts><Host Name="Document"/></Hosts>
	<DefaultSettings><SourceLocation DefaultValue="javascript:void(0)"/></DefaultSettings>
</OfficeApp>`;

export const cases: readonly SecurityCase[] = [
	{
		name: 'executable task pane URLs are rejected',
		limits: { ms: 2_000, heapMb: 64 },
		run: () => refusal(new TextEncoder().encode(unsafePane)),
		expected: 'url',
	},
	{
		name: 'oversized detached manifests are rejected',
		limits: { ms: 2_000, heapMb: 64 },
		run: () => refusal(new Uint8Array(256 * 1024 + 1).fill(0x20)),
		expected: 'limit',
	},
	{
		name: 'a truncated manifest archive is rejected',
		limits: { ms: 2_000, heapMb: 64 },
		run: () => refusal(new Uint8Array([0x50, 0x4b, 3, 4])),
		expected: 'package',
	},
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
			return refusal(await zip.generateAsync({ type: 'uint8array' }), 'field');
		},
		expected: 'UNSAFE_PATH',
	},
	{
		name: 'archive entries must stay within the compression ratio limit',
		limits: { ms: 2_000, heapMb: 64 },
		async run() {
			const zip = new JSZip().file('manifest.json', '{}').file('padding', new Uint8Array(16_384));
			const data = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
			return refusal(data, 'field');
		},
		expected: 'LIMIT_RATIO',
	},
];
