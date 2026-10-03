import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { createRequire } from 'node:module';
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';
import { build } from 'vite';
import { VIEWER_PACKAGES } from './release-plan.mjs';
import { verifyManifest } from './publish-released.mjs';
import { runNpm } from './npm-command.mjs';
import { createVsdxFixture } from '../tests/fixture.mjs';

const root = resolve(import.meta.dirname, '..');
const consumer = mkdtempSync(resolve(tmpdir(), 'visio-published-consumer-'));
const dependencies = {};
for (const [key, meta] of Object.entries(VIEWER_PACKAGES)) {
	const directory = resolve(root, meta.dir);
	const manifest = JSON.parse(readFileSync(resolve(directory, 'package.json')));
	verifyManifest({ ...meta, version: manifest.version });
	const [pack] = JSON.parse(
		runNpm(['pack', '--json', '--ignore-scripts', '--pack-destination', consumer], {
			cwd: directory,
			encoding: 'utf8',
		}),
	);
	const files = pack.files.map((file) => file.path);
	for (const file of ['dist/index.js', 'dist/index.d.ts', 'LICENSE'])
		assert.ok(files.includes(file), `${meta.npm}: ${file}`);
	assert.ok(
		!files.some((file) => /\.test\.|integration\/|node_modules\//.test(file)),
		'No development artifacts',
	);
	if (key !== 'core') {
		for (const worker of ['parse-worker-', 'edit-worker-'])
			assert.ok(
				files.some((file) => file.startsWith('dist/assets/' + worker)),
				`${meta.npm}: ${worker}`,
			);
	}
	dependencies[manifest.name] = `file:${resolve(consumer, pack.filename)}`;
	Object.assign(dependencies, manifest.peerDependencies);
}
// Every sibling must resolve to its tarball even when dependent manifests use registry ranges.
writeFileSync(
	resolve(consumer, 'package.json'),
	JSON.stringify({ name: 'visio-package-consumer', private: true, type: 'module', dependencies }),
);
runNpm(['install', '--ignore-scripts'], { cwd: consumer, stdio: 'inherit' });
writeFileSync(resolve(consumer, 'fixture.vsdx'), await createVsdxFixture('Published consumer'));
const names = Object.values(VIEWER_PACKAGES).map((meta) => meta.npm);
writeFileSync(
	resolve(consumer, 'smoke.mjs'),
	`
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
await import('@angular/compiler');
const bytes = readFileSync(new URL('./fixture.vsdx', import.meta.url));
for (const name of ${JSON.stringify(names)}) {
 const api = await import(name);
 const document = await api.parseVsdx(bytes);
 assert.equal(document.pages[0].shapes[0].text.plainText, 'Published consumer');
 if (name !== 'visio-core') {
  assert.equal(typeof api.mountViewer, 'function', name);
  const controller = new api.ViewerController();
  controller.setZoom(1.25);
  assert.equal(controller.state.zoom, 1.25);
 }
}
`,
);
// Angular's compiler is a host concern for JIT applications, not a bundled viewer dependency.
runNpm(['install', '--ignore-scripts', '--no-save', '@angular/compiler@^21'], {
	cwd: consumer,
	stdio: 'inherit',
});
execFileSync(process.execPath, [resolve(consumer, 'smoke.mjs')], {
	cwd: consumer,
	stdio: 'inherit',
});
writeFileSync(
	resolve(consumer, 'types.ts'),
	names
		.map(
			(name, index) =>
				`import * as p${index} from '${name}';\nvoid p${index}.parseVsdx;${index ? `void new p${index}.ViewerController();` : ''}`,
		)
		.join('\n'),
);
execFileSync(
	process.execPath,
	[
		resolve(root, 'node_modules/typescript/bin/tsc'),
		'--noEmit',
		'--module',
		'nodenext',
		'--target',
		'es2022',
		'--lib',
		'es2022,dom',
		'--skipLibCheck',
		'types.ts',
	],
	{ cwd: consumer, stdio: 'inherit' },
);
writeFileSync(
	resolve(consumer, 'index.html'),
	'<input type="file" id="file"><div id="app" style="height:600px"></div><script type="module" src="./main.js"></script>',
);
writeFileSync(
	resolve(consumer, 'main.js'),
	`import { mountViewer } from 'visio-vanilla-viewer';
import '@angular/compiler';
import { VisioViewer as ReactViewer } from 'visio-react-viewer';
import { VisioViewer as VueViewer } from 'visio-vue-viewer';
import { VisioViewer as SolidViewer } from 'visio-solid-viewer';
import { VisioViewerComponent } from 'visio-angular-viewer';
import { VisioViewer as SvelteViewer, parseVsdx } from 'visio-svelte-viewer';
window.adapters = { ReactViewer, VueViewer, SolidViewer, VisioViewerComponent, SvelteViewer, parseVsdx };
window.viewer = mountViewer(document.getElementById('app'));
document.getElementById('file').addEventListener('change', async event => {
 try { await window.viewer.load(event.target.files[0]); } catch (error) { window.loadError = String(error); }
});`,
);
const bindingsRequire = createRequire(resolve(root, 'packages/bindings/package.json'));
const { svelte } = await import(
	pathToFileURL(bindingsRequire.resolve('@sveltejs/vite-plugin-svelte')).href
);
await build({ configFile: false, root: consumer, plugins: [svelte()], build: { outDir: 'dist' } });
writeFileSync(resolve(root, '.package-build/consumer.txt'), consumer);
execFileSync(
	process.execPath,
	[resolve(root, 'scripts/test-worker-bundle.mjs'), resolve(consumer, 'dist/assets')],
	{ cwd: root, stdio: 'inherit' },
);
console.log(
	'Seven tarballs pass registry-only install, ESM imports, VSDX parsing, declarations and consumer worker checks.',
);
