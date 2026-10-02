import { execFileSync } from 'node:child_process';
import { mkdtempSync, writeFileSync, readdirSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { resolve, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import assert from 'node:assert/strict';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '..'),
	temporary = mkdtempSync(resolve(tmpdir(), 'visio-packed-consumer-'));
const env = {
	...process.env,
	npm_config_cache: resolve(temporary, 'npm-cache'),
};
const run = (command, args, cwd = temporary) =>
	execFileSync(command, args, { cwd, env, stdio: 'inherit' });
const packed = JSON.parse(
	execFileSync('npm', ['pack', '--json', '--ignore-scripts', '--pack-destination', temporary], {
		cwd: root,
		env,
		encoding: 'utf8',
	}),
)[0];
assert.ok(
	packed.files.some((file) => file.path === 'dist/parse-worker.js'),
	'Worker entry must ship in the actual npm artifact',
);
assert.ok(
	!packed.files.some((file) => /\.test\.|test-fixture|integration\//.test(file.path)),
	'Fixtures and development patches must not ship in the package',
);
writeFileSync(
	resolve(temporary, 'package.json'),
	JSON.stringify({
		name: 'visio-packed-consumer',
		version: '0.0.0',
		private: true,
		type: 'module',
		dependencies: {
			'ooxml-core': `file:${resolve(root, '..', 'ooxml')}`,
			'@christophervr/visio-viewer': `file:${resolve(temporary, packed.filename)}`,
		},
	}),
);
run('npm', ['install', '--ignore-scripts']);
run(process.execPath, [
	'--input-type=module',
	'-e',
	`const m=await import('@christophervr/visio-viewer');if(!m.ViewerController||!m.registerVisioViewer||!m.exportPageSvg||m.MAX_SVG_EXPORT_BYTES!==16777216||m.TEXT_SEARCH_LIMITS.queryCharacters!==256)throw Error('Missing public API');`,
]);
writeFileSync(
	resolve(temporary, 'smoke.ts'),
	`import { ViewerController, exportPageSvg, type SvgExportOptions, type SvgExportResult, type MountedViewer, type ViewerOptions, type TextSearchState } from '@christophervr/visio-viewer'; const controller=new ViewerController(); const props:ViewerOptions={zoom:1}; controller.setZoom(props.zoom!); controller.setSearchQuery("literal text"); controller.nextSearchResult(); const search:TextSearchState=controller.state.search; void search; const exportOptions:SvgExportOptions={maxBytes:100000}; const exporter:(...args:Parameters<typeof exportPageSvg>)=>SvgExportResult=exportPageSvg; function snapshot(viewer:MountedViewer):SvgExportResult { return viewer.exportSvg(exportOptions); } void exporter; void snapshot;`,
);
run(process.execPath, [
	resolve(root, 'node_modules/typescript/bin/tsc'),
	'--noEmit',
	'--module',
	'nodenext',
	'--target',
	'es2022',
	'--lib',
	'es2022,dom',
	'--skipLibCheck',
	'smoke.ts',
]);
writeFileSync(
	resolve(temporary, 'index.html'),
	'<div id="app"></div><script type="module" src="./main.js"></script>',
);
writeFileSync(
	resolve(temporary, 'main.js'),
	`import { registerVisioViewer } from '@christophervr/visio-viewer'; registerVisioViewer();document.getElementById('app').append(document.createElement('visio-viewer'));`,
);
writeFileSync(
	resolve(temporary, 'vite.config.mjs'),
	`export default {base:'./',build:{outDir:'build'}};`,
);
run(process.execPath, [resolve(root, 'node_modules/vite/bin/vite.js'), 'build']);
assert.ok(
	readdirSync(resolve(temporary, 'build/assets')).some((file) =>
		/^parse-worker-.*\.js$/.test(file),
	),
	'Consumer build must resolve the packaged worker',
);
run(process.execPath, [
	resolve(root, 'scripts/test-worker-bundle.mjs'),
	resolve(temporary, 'build/assets'),
]);
console.log(`Packed ESM/type declarations/consumer build/worker checks passed: ${temporary}`);
