import { readFileSync } from 'node:fs';
import JSZip from 'jszip';
import { describe, expect, it } from 'vitest';
import { MANIFEST_BYTES, PACKAGE_BYTES, readLocalAddIn, type AddInReadOptions } from './index';

const fixture = (name: string) =>
	readFileSync(new URL(`./__fixtures__/${name}`, import.meta.url), 'utf8');
const xml = fixture('manifest.xml');
const unified = fixture('manifest.json');
const bytes = (source: string) => new TextEncoder().encode(source);
const word: AddInReadOptions = { host: 'word' };
const support: AddInReadOptions = {
	...word,
	capabilities: { sets: new Map([['addincommands', '1.10']]) },
};
const read = (source: string, options = word) => readLocalAddIn(bytes(source), options);
const firstCommand = (manifest: Awaited<ReturnType<typeof read>>['manifest']) =>
	manifest.tabs[0]?.groups[0]?.commands[0];

async function pack(source = unified, add?: (zip: JSZip) => void) {
	const zip = new JSZip()
		.file('manifest.json', source)
		.file('icons/outline.png', new Uint8Array([1]))
		.file('icons/color.png', new Uint8Array([2]))
		.file('fr.json', fixture('fr.json'));
	add?.(zip);
	return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

describe('local Office add-in manifests', () => {
	it('resolves the same identity, permissions and pane from independently authored XML and JSON', async () => {
		const { manifest: a, digest } = await read(xml, support);
		const { manifest: b } = await readLocalAddIn(await pack(), support);
		for (const key of [
			'id',
			'version',
			'publisher',
			'name',
			'description',
			'hosts',
			'permissions',
			'declaredOrigins',
			'pane',
		] as const)
			expect(b[key]).toEqual(a[key]);
		expect(firstCommand(b)).toEqual(firstCommand(a));
		expect(a.tabs[0]?.groups[0]?.commands[1]).toMatchObject({
			kind: 'inactive',
			issue: { feature: 'ExecuteFunction' },
		});
		expect(b.icons).toEqual([
			{ kind: 'package', path: 'icons/color.png' },
			{ kind: 'package', path: 'icons/outline.png' },
		]);
		expect(digest).toMatch(/^[\da-f]{64}$/);
		expect((await read(xml)).digest).toBe(digest);
		expect((await read(xml.replace('1.0.0', '1.0.1'))).digest).not.toBe(digest);
	});

	it('uses exact locale, parent, then manifest default for text and URLs', async () => {
		const xmlFr = await read(xml, { ...support, locale: 'fr-CA' });
		expect(xmlFr.manifest.name).toBe('Carnet de labo');
		expect(xmlFr.manifest.pane).toBe('https://notes.example.test/fr/pane');
		expect(firstCommand(xmlFr.manifest)).toHaveProperty(
			'url',
			'https://notes.example.test/fr/pane',
		);
		expect((await read(xml, { ...support, locale: 'de' })).manifest.name).toBe('Notes');
		const { manifest } = await readLocalAddIn(await pack(), { ...support, locale: 'fr-CA' });
		expect(manifest.name).toBe('Notes de labo');
		expect(firstCommand(manifest)?.label).toBe('Ouvrir les notes');
	});

	it('checks base XML requirements before overrides and falls back by numeric version', async () => {
		const { manifest } = await read(xml, {
			...word,
			capabilities: { sets: new Map([['addincommands', '1.2']]) },
		});
		expect(manifest.pane).toBe('https://notes.example.test/pane');
		expect(firstCommand(manifest)).toMatchObject({
			kind: 'inactive',
			issue: { requirement: { minVersion: '1.10' } },
		});
		const base = xml.replace(
			'<DefaultSettings>',
			'<Requirements><Sets><Set Name="WordApi" MinVersion="1.1" /></Sets></Requirements><DefaultSettings>',
		);
		const unsupported = await read(base, support);
		expect(unsupported.manifest.pane).toBeNull();
		expect(unsupported.manifest.tabs).toEqual([]);
		expect(unsupported.manifest.issues).toContainEqual({
			kind: 'requirement',
			requirement: { kind: 'set', name: 'WordApi', minVersion: '1.1' },
		});
	});

	it.each([
		['broken markup', xml.replace('</OfficeApp>', '</Office>'), 'manifest'],
		[
			'missing xsi namespace',
			xml.replace('http://www.w3.org/2001/XMLSchema-instance', 'urn:fake'),
			'manifest',
		],
		[
			'DTD',
			xml.replace(
				'<OfficeApp',
				'<!DOCTYPE OfficeApp SYSTEM "https://hostile.example.test/entity">\n<OfficeApp',
			),
			'manifest',
		],
		[
			'executable locale URL',
			xml.replace('https://notes.example.test/fr/pane', 'javascript:alert(1)'),
			'url',
		],
		[
			'credentials',
			xml.replace('https://notes.example.test/pane', 'https://user:pass@notes.example.test'),
			'url',
		],
		[
			'external override origin',
			xml.replace(
				'DefaultValue="https://notes.example.test/pane">\n\t\t\t\t\t',
				'DefaultValue="https://external.example.test/pane">\n\t\t\t\t\t',
			),
			'url',
		],
		['unresolved resource', xml.replace('resid="Notes.Open"', 'resid="Missing"'), 'manifest'],
		['duplicate resource', xml.replace('id="Notes.Function"', 'id="notes.open"'), 'manifest'],
		[
			'ambiguous permissions',
			xml.replace('</Permissions>', '</Permissions><Permissions>Restricted</Permissions>'),
			'manifest',
		],
		['other host', xml.replace('<Host Name="Document" />', ''), 'host'],
		['content extension', xml.replace('TaskPaneApp', 'ContentApp'), 'extension'],
	])('rejects XML with %s', async (_case, input, code) => {
		await expect(read(input)).rejects.toMatchObject({ code });
	});

	it('checks unified requirements at the extension, runtime and ribbon and reports other features', async () => {
		const source = unified
			.replace(
				'"extensions": [',
				'"bots": [{"botId":"inactive"}], "extensions": [{"requirements":{"scopes":["mail"]}},',
			)
			.replace(
				'"contexts": ["default"]',
				'"requirements":{"capabilities":[{"name":"WordApi","maxVersion":"1.2"}]},"contexts":["default"]',
			);
		const denied = await readLocalAddIn(await pack(source), support);
		expect(denied.manifest.issues).toContainEqual({ kind: 'unsupported', feature: 'bots' });
		expect(denied.manifest.issues).toContainEqual({ kind: 'unsupported', feature: 'mail' });
		expect(firstCommand(denied.manifest)).toMatchObject({
			kind: 'inactive',
			issue: { requirement: { name: 'WordApi' } },
		});
		const atMaximum = {
			...word,
			capabilities: {
				sets: new Map([
					['addincommands', '1.10'],
					['wordapi', '1.2'],
				]),
			},
		};
		expect(firstCommand((await readLocalAddIn(await pack(source), atMaximum)).manifest)?.kind).toBe(
			'taskPane',
		);
		const extensionDenied = unified.replace(
			'"scopes": ["document"',
			'"capabilities":[{"name":"WordApi"}],"scopes":["document"',
		);
		expect(
			firstCommand((await readLocalAddIn(await pack(extensionDenied), support)).manifest)?.kind,
		).toBe('inactive');
		expect(firstCommand((await readLocalAddIn(await pack(), word)).manifest)?.kind).toBe(
			'inactive',
		);
	});

	it.each([
		[
			'duplicate escaped key',
			unified.replace('"version":', '"ver\\u0073ion": "1.0.1", "version":'),
			'manifest',
		],
		['unknown version', unified.replace('"1.25"', '"9.99"'), 'version'],
		[
			'duplicate action ID',
			unified.replace('"actions": [', '"actions": [{"id":"opennotes","type":"openPage"},'),
			'manifest',
		],
		[
			'unresolved action',
			unified.replace('"actionId": "OpenNotes"', '"actionId": "Missing"'),
			'manifest',
		],
		['unsupported host mapping', unified.replace('"document"', '"anything"'), 'manifest'],
		['wildcard origin', unified.replace('assets.example.test', '*.example.test'), 'url'],
		['script URL', unified.replace('https://notes.example.test/pane', 'data:text/html,bad'), 'url'],
		[
			'deep JSON',
			unified.replace('"name":', `"deep":${'['.repeat(33)}0${']'.repeat(33)},"name":`),
			'manifest',
		],
	])('rejects unified JSON with %s', async (_case, input, code) => {
		await expect(readLocalAddIn(await pack(input), support)).rejects.toMatchObject({ code });
	});

	it('checks manifest and ZIP byte, path, count and inflate limits', async () => {
		await expect(readLocalAddIn(new Uint8Array(PACKAGE_BYTES + 1), word)).rejects.toMatchObject({
			code: 'limit',
		});
		await expect(read(xml.padEnd(MANIFEST_BYTES + 1))).rejects.toMatchObject({ code: 'limit' });
		await expect(
			readLocalAddIn(
				await pack(unified, (zip) => {
					for (let i = 0; i < 128; i++) zip.file(`file-${i}`, 'a');
				}),
				word,
			),
		).rejects.toMatchObject({ code: 'limit' });
		await expect(
			readLocalAddIn(
				await pack(unified, (zip) => {
					zip.file('large', new Uint8Array(PACKAGE_BYTES));
				}),
				word,
			),
		).rejects.toMatchObject({ code: 'limit' });
		for (const path of ['../outside', '%2e%2e/outside', '/absolute', 'a\\b', 'MANIFEST.JSON']) {
			await expect(
				readLocalAddIn(
					await pack(unified, (zip) => {
						zip.file(path, 'bad');
					}),
					word,
				),
			).rejects.toMatchObject({ code: 'package' });
		}
	});

	it('validates icon and locale references and permits a detached JSON file without package assets', async () => {
		await expect(
			readLocalAddIn(
				await pack(unified, (zip) => {
					zip.remove('icons/color.png');
				}),
				word,
			),
		).rejects.toMatchObject({ code: 'package' });
		await expect(
			readLocalAddIn(
				await pack(unified, (zip) => {
					zip.file('fr.json', '{"name.short":"x","name.short":"y"}');
				}),
				word,
			),
		).rejects.toMatchObject({ code: 'manifest' });
		const detached = unified
			.replace(/"icons": [^\n]+\n/, '')
			.replace(/"additionalLanguages": \[[^\n]+/, '"additionalLanguages": []');
		expect((await read(detached)).manifest.format).toBe('unified');
	});
});
