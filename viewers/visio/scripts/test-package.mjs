import { runNpm } from './npm-command.mjs';
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
	runNpm(['pack', '--json', '--ignore-scripts', '--pack-destination', temporary], {
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
	packed.files.some((file) => file.path === 'dist/edit-worker.js'),
	'Edit worker entry must ship in the actual npm artifact',
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
			'ooxml-core': '0.12.0',
			'@christophervr/visio-viewer': `file:${resolve(temporary, packed.filename)}`,
		},
	}),
);
runNpm(['install', '--ignore-scripts'], { cwd: temporary, env, stdio: 'inherit' });
run(process.execPath, [
	'--input-type=module',
	'-e',
	`const m=await import('@christophervr/visio-viewer');if(!m.ViewerController||!m.registerVisioViewer||!m.exportPageSvg||m.MAX_SVG_EXPORT_BYTES!==16777216||m.TEXT_SEARCH_LIMITS.queryCharacters!==256||!m.createPrintSnapshot||m.PRINT_SNAPSHOT_LIMITS.maxPages!==32||m.VIEWER_LAYER_LIMITS.controls!==200)throw Error('Missing public API');`,
]);
writeFileSync(
	resolve(temporary, 'smoke.ts'),
	`import { ViewerController, exportPageSvg, type SvgExportOptions, type SvgExportResult, type MountedViewer, type ViewerOptions, type TextSearchState, type PrintSnapshot, type CurrentPagePrintSnapshotOptions, type LayerVisibilityOverride, type ViewerEditState, type VsdxExportResult } from '@christophervr/visio-viewer'; const controller=new ViewerController(); const props:ViewerOptions={zoom:1}; controller.setZoom(props.zoom!); controller.setSearchQuery("literal text"); controller.nextSearchResult(); const search:TextSearchState=controller.state.search; void search; const exportOptions:SvgExportOptions={maxBytes:100000}; const exporter:(...args:Parameters<typeof exportPageSvg>)=>SvgExportResult=exportPageSvg; function snapshot(viewer:MountedViewer):SvgExportResult { return viewer.exportSvg(exportOptions); } const printOptions:CurrentPagePrintSnapshotOptions={limits:{maxPages:1}}; function prepare(viewer:MountedViewer):PrintSnapshot { return viewer.createPrintSnapshot(printOptions); } function layers(viewer:MountedViewer):void { viewer.setLayerVisibility("0","1",false); viewer.setLayerVisibility("0","1",null); viewer.resetLayerVisibility("0"); viewer.resetLayerVisibility(); } const overrides:readonly LayerVisibilityOverride[]=controller.state.layerVisibilityOverrides; void overrides; void layers; void exporter; void snapshot; void prepare; const editState:ViewerEditState=controller.state.edit; const saved:VsdxExportResult=controller.exportVsdx(); const edit:Promise<void>=controller.replacePlainText('1','1','literal'); const undo:Promise<void>=controller.undo(); const redo:Promise<void>=controller.redo(); controller.cancelEdit(); function mountedEdit(viewer:MountedViewer):VsdxExportResult { void viewer.replacePlainText('1','1','literal'); void viewer.undo(); void viewer.redo(); viewer.cancelEdit(); return viewer.exportVsdx(); } void editState; void saved; void edit; void undo; void redo; void mountedEdit;`,
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
assert.ok(
	readdirSync(resolve(temporary, 'build/assets')).some((file) => /^edit-worker-.*\.js$/.test(file)),
	'Consumer build must resolve the packaged edit worker',
);
run(process.execPath, [
	resolve(root, 'scripts/test-worker-bundle.mjs'),
	resolve(temporary, 'build/assets'),
]);
console.log(`Packed ESM/type declarations/consumer build/worker checks passed: ${temporary}`);
