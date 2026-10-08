// Packs the built UI package AND the core package, installs both tarballs into a clean project
// and imports every entry point the way a consumer would: once under plain Node (SSR: must import
// and register as a no-op) and once inside a jsdom window (elements actually register).
// Run after `bun run build` in src/core and in src/ui.
import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { mkdtemp, mkdir, readFile, rm, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const uiRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const coreRoot = path.resolve(uiRoot, '..', 'core');
const manifest = JSON.parse(await readFile(path.join(uiRoot, 'package.json'), 'utf8'));
const work = await mkdtemp(path.join(tmpdir(), 'office-ui-smoke-'));
const npmCli =
	process.platform === 'win32'
		? path.join(path.dirname(process.execPath), 'node_modules', 'npm', 'bin', 'npm-cli.js')
		: undefined;
const run = (command, args, cwd) => {
	const executable = command === 'npm' && npmCli ? process.execPath : command;
	const executableArgs = command === 'npm' && npmCli ? [npmCli, ...args] : args;
	const result = spawnSync(executable, executableArgs, { cwd, encoding: 'utf8' });
	if (result.status !== 0)
		throw new Error(`${command} ${args.join(' ')} failed\n${result.stdout}\n${result.stderr}`);
	return result.stdout;
};
const pack = (cwd) =>
	JSON.parse(
		run('npm', ['pack', '--json', '--ignore-scripts', '--pack-destination', work], cwd),
	)[0];

try {
	const ui = pack(uiRoot);
	const core = pack(coreRoot);
	const files = new Set(ui.files.map((entry) => entry.path));
	for (const required of [
		'LICENSE',
		'NOTICE',
		'README.md',
		'package.json',
		'dist/index.js',
		'dist/index.d.ts',
	])
		assert(files.has(required), `the package is missing ${required}`);
	assert(
		![...files].some((file) => /\.test\.|__tests__|__fixtures__|src\//.test(file)),
		'tests or sources were packed',
	);
	assert(manifest.dependencies['ooxml-core'], 'the core dependency is missing');
	assert(
		!JSON.stringify(manifest.dependencies).match(/file:|workspace:|link:/),
		'a runtime dependency uses file:, workspace: or link:',
	);
	assert(
		Object.keys(manifest.peerDependencies ?? {}).every(
			(name) => manifest.peerDependenciesMeta?.[name]?.optional,
		),
		'product integrations must remain optional',
	);

	await writeFile(
		path.join(work, 'package.json'),
		JSON.stringify({ private: true, type: 'module' }),
	);
	run(
		'npm',
		[
			'install',
			'--ignore-scripts',
			'--no-audit',
			'--no-fund',
			'--package-lock=false',
			'jsdom@^26.1.0',
			path.join(work, core.filename),
			path.join(work, ui.filename),
		],
		work,
	);

	// Data files (package.json, the Custom Elements Manifest) are not modules; they are read below.
	const entries = Object.keys(manifest.exports)
		.filter((entry) => !entry.endsWith('.json'))
		.flatMap((entry) => {
			if (!entry.includes('*')) return [entry];
			const [prefix, suffix] = manifest.exports[entry].import.slice(2).split('*');
			return [...files]
				.filter((file) => file.startsWith(prefix) && file.endsWith(suffix))
				.map((file) => entry.replaceAll('*', file.slice(prefix.length, -suffix.length)));
		});
	const specifiers = entries.map((entry) =>
		path.posix.join(manifest.name, entry === '.' ? '' : entry),
	);
	await mkdir(path.join(work, 'check'), { recursive: true });
	// 1. Server-side: importing every entry needs no DOM and registering does nothing.
	await writeFile(
		path.join(work, 'ssr.mjs'),
		`import assert from 'node:assert/strict';
const specifiers = ${JSON.stringify(specifiers)};
for (const s of specifiers) assert.ok(Object.keys(await import(s)).length > 0, s + ' has no exports');
import { createRequire } from 'node:module';
const cem = createRequire(import.meta.url)('${manifest.name}/custom-elements.json');
assert.ok(cem.modules.length > 0, 'the Custom Elements Manifest is empty');
const { registerOfficeUi } = await import('${manifest.name}');
registerOfficeUi();
console.log('ssr import ok: ' + specifiers.length + ' entries');
`,
	);
	// 2. In a DOM: the globals are set before the (single) import, then every tag must exist.
	await writeFile(
		path.join(work, 'dom.mjs'),
		`import assert from 'node:assert/strict';
import { JSDOM } from 'jsdom';
const dom = new JSDOM('<!doctype html><body></body>', { pretendToBeVisual: true });
for (const key of ['window', 'document', 'HTMLElement', 'customElements', 'CustomEvent', 'Event', 'KeyboardEvent', 'Node', 'Element'])
	Object.defineProperty(globalThis, key, { value: dom.window[key], configurable: true, writable: true });
const ui = await import('${manifest.name}');
ui.registerOfficeUi();
ui.registerOfficeUi();
for (const tag of ui.OFFICE_UI_TAGS) assert.ok(customElements.get(tag), tag + ' is not defined');
const pptx = await import('${manifest.name}/pptx');
pptx.registerPptxWebControls();
pptx.registerPptxWebControls();
assert.ok(customElements.get('pptx-ui-title-bar'), 'the PowerPoint title bar is not defined');
assert.ok(customElements.get('pptx-ui-search'), 'the PowerPoint search alias is not defined');
assert.equal(typeof pptx.canDrillDown, 'function');
assert.ok(pptx.vermilionDarkTheme.colors);
const smartart = document.createElement('office-ui-smartart');
document.body.append(smartart);
smartart.drawing = { issues: [], shapes: [{ modelId: '1', frame: { x: 0, y: 0, width: 952500, height: 476250 }, geometry: 'rect', has3d: false,
	fill: { kind: 'solid', color: { kind: 'srgb', value: '4472C4', transforms: [] } } }] };
assert.equal(smartart.shadowRoot.querySelectorAll('rect').length, 1);
console.log('dom registration ok: ' + ui.OFFICE_UI_TAGS.length + ' tags');
`,
	);
	console.log(run('node', ['ssr.mjs'], work).trim());
	console.log(run('node', ['dom.mjs'], work).trim());
} finally {
	await rm(work, { recursive: true, force: true });
}
