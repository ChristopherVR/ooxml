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
		const unsupported = await read(base);
		expect(unsupported.manifest.pane).toBeNull();
		expect(unsupported.manifest.tabs).toEqual([]);
		expect(unsupported.manifest.issues).toEqual([
			{ kind: 'requirement', requirement: { kind: 'set', name: 'WordApi', minVersion: '1.1' } },
			{
				kind: 'requirement',
				requirement: { kind: 'set', name: 'AddinCommands', minVersion: '1.10' },
			},
			{ kind: 'unsupported', feature: 'ExecuteFunction' },
		]);
	});

	it('accepts schema hosts outside the editor targets and ignores their overrides', async () => {
		const extra = ['Project', 'Notebook', 'Database', 'Mailbox']
			.map((name) => `<Host Name="${name}" />`)
			.join('');
		const source = xml
			.replace('<Hosts><Host Name="Document" />', `<Hosts>${extra}<Host Name="Document" />`)
			.replace(
				'<Host xsi:type="Document">',
				'<Host xsi:type="Notebook" /><Host xsi:type="Document">',
			);
		const { manifest } = await read(source, support);
		expect(manifest.hosts).toEqual(['word', 'excel', 'powerpoint']);
		expect(manifest.tabs).toHaveLength(1);
		await expect(
			read(xml.replace(/<Hosts>.*?<\/Hosts>/, `<Hosts>${extra}</Hosts>`)),
		).rejects.toMatchObject({
			code: 'host',
		});
		await expect(read(source.replace('Name="Project"', 'Name="Unknown"'))).rejects.toMatchObject({
			code: 'manifest',
			field: 'host name',
		});
		await expect(
			read(source.replace('xsi:type="Notebook"', 'xsi:type="Project"')),
		).rejects.toMatchObject({
			code: 'manifest',
			field: 'override host type',
		});
	});

	it('reports XML override issues for the selected host', async () => {
		const excelHost = `<Host xsi:type="Workbook"><DesktopFormFactor>
			<ExtensionPoint xsi:type="ContextMenu" />
			<ExtensionPoint xsi:type="PrimaryCommandSurface"><OfficeTab id="TabHome">
				<Group id="NotesGroup"><Label resid="Notes.Group" />
					<Control xsi:type="Button" id="OpenNotes"><Label resid="Notes.Open" />
						<Action xsi:type="ShowTaskpane">
							<TaskpaneId>Office.AutoShowTaskpaneWithDocument</TaskpaneId>
							<SourceLocation resid="Notes.Pane" />
						</Action>
					</Control>
					<Control xsi:type="ToggleButton" id="ToggleNotes" />
				</Group>
			</OfficeTab></ExtensionPoint>
		</DesktopFormFactor></Host>`;
		const source = xml.replace('</Host>', `</Host>${excelHost}`);
		const { manifest: wordManifest } = await read(source, support);
		expect(wordManifest.issues).toEqual([{ kind: 'unsupported', feature: 'ExecuteFunction' }]);
		const { manifest: excel } = await read(source, { host: 'excel' });
		expect(excel.issues).toEqual([
			{
				kind: 'requirement',
				requirement: { kind: 'set', name: 'AddinCommands', minVersion: '1.10' },
			},
			...['ContextMenu', 'built-in ribbon tab', 'automatic task pane', 'ToggleButton'].map(
				(feature) => ({ kind: 'unsupported', feature }),
			),
		]);
		const { manifest: powerpoint } = await read(source, { host: 'powerpoint' });
		expect(powerpoint.tabs).toEqual([]);
		expect(powerpoint.issues).toEqual([]);
		await expect(read(source.replace('id="TabHome"', 'id=""'), support)).rejects.toMatchObject({
			code: 'manifest',
			field: 'tab ID',
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
		expect(denied.manifest.issues).not.toContainEqual({ kind: 'unsupported', feature: 'mail' });
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

	it('reports unified runtime and ribbon issues for the selected host', async () => {
		const excelExtension = {
			requirements: { scopes: ['workbook'] },
			runtimes: [
				{
					id: 'ExcelRuntime',
					type: 'general',
					lifetime: 'long',
					code: { page: 'https://notes.example.test/excel' },
					actions: [{ id: 'Execute', type: 'executeFunction' }],
				},
			],
			ribbons: [{ contexts: ['contextMenu'] }],
		};
		const source = unified
			.replace('"document", "workbook", "presentation"', '"document"')
			.replace('"extensions": [', `"extensions": [${JSON.stringify(excelExtension)},`);
		const archive = await pack(source);
		const { manifest: wordManifest } = await readLocalAddIn(archive, support);
		expect(wordManifest.issues).toEqual([]);
		expect(firstCommand(wordManifest)?.kind).toBe('taskPane');
		const { manifest: excel } = await readLocalAddIn(archive, { host: 'excel' });
		expect(excel.issues).toEqual(
			['shared runtime', 'executeFunction', 'contextMenu'].map((feature) => ({
				kind: 'unsupported',
				feature,
			})),
		);
		const invalidExcel = await pack(source.replace('"lifetime":"long"', '"lifetime":"invalid"'));
		await expect(readLocalAddIn(invalidExcel, support)).rejects.toMatchObject({
			code: 'manifest',
			field: 'runtime lifetime',
		});
	});

	it.each(['word', 'excel', 'powerpoint'] as const)(
		'reports all applicable runtime issues in %s',
		async (host) => {
			const source = unified
				.replace('"lifetime": "short"', '"lifetime": "long"')
				.replace(
					'"minVersion": "1.10" }',
					'"minVersion": "1.10" }, { "name": "SharedRuntime", "minVersion": "1.1" }',
				);
			const { manifest } = await readLocalAddIn(await pack(source), { host });
			expect(manifest.issues).toEqual([
				{
					kind: 'requirement',
					requirement: { kind: 'set', name: 'AddinCommands', minVersion: '1.10' },
				},
				{
					kind: 'requirement',
					requirement: { kind: 'set', name: 'SharedRuntime', minVersion: '1.1' },
				},
				{ kind: 'unsupported', feature: 'shared runtime' },
			]);
		},
	);

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
		const original = (await read(unified, { ...support, locale: 'fr' })).manifest;
		expect(original.name).toBe('Notes');
		expect(original.icons).toEqual([
			{ kind: 'package', path: 'icons/color.png' },
			{ kind: 'package', path: 'icons/outline.png' },
		]);
	});

	it('rejects top-level XML elements in the wrong order or namespace', async () => {
		await expect(
			read(xml.replace(/(<Version>.*?<\/Version>)\s*(<ProviderName>.*?<\/ProviderName>)/, '$2$1')),
		).rejects.toMatchObject({ code: 'manifest', field: 'Version' });
		await expect(read(xml.replace('<Hosts>', '<Hosts xmlns="urn:fake">'))).rejects.toMatchObject({
			code: 'manifest',
			field: 'Hosts',
		});
	});
});
