/**
 * Proof that `pptx-react-viewer/no-i18next` works without i18next and react-i18next, run by
 * `check-binding-packages.mjs` on the packed packages:
 *
 * - no file reachable from its ESM or CommonJS build imports either library, while the main
 *   entry's build still does;
 * - a consumer that installs the packed packages without them (and checks they are absent)
 *   renders the entry with Node in both module formats, and bundles an app importing it for the
 *   browser, while the same app importing `pptx-react-viewer` fails to resolve react-i18next.
 */
import assert from 'node:assert/strict';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';

import { bundleBrowserApp } from '../../../scripts/browser-bundle-guard.mjs';
import { walkModuleGraph } from '../../../scripts/module-graph.mjs';

const NAME = 'pptx-react-viewer';
const ENTRY = `${NAME}/no-i18next`;
const I18N_PEERS = ['i18next', 'react-i18next'];
const I18NEXT = /^(?:react-)?i18next(?:\/|$)/u;
const viewerRoot = resolve(import.meta.dirname, '..');
const manifests = ['packages/react', '../../src/ui', '../../src/core'].map((dir) =>
	JSON.parse(readFileSync(join(viewerRoot, dir, 'package.json'), 'utf8')),
);

/** Every i18next import reachable from an installed file, lazily loaded chunks included. */
async function i18nextImports(file) {
	const { bare } = await walkModuleGraph(file);
	return bare
		.filter(({ specifier }) => I18NEXT.test(specifier))
		.map(({ specifier, kind, importer }) => `${kind} of ${specifier} in ${importer}`);
}

/**
 * The built graphs of `./no-i18next` reach neither library; the main entry still does. Both
 * entries ship the same declarations, so their types are the same.
 */
export async function assertNoI18nextInBuiltGraph(packageDir) {
	const { exports } = JSON.parse(readFileSync(join(packageDir, 'package.json'), 'utf8'));
	const types = (subpath) => readFileSync(join(packageDir, exports[subpath].types), 'utf8');
	assert.equal(types('./no-i18next'), types('.'), `${ENTRY} declares the types of ${NAME}`);
	for (const condition of ['import', 'require']) {
		const file = exports['./no-i18next'][condition];
		assert.deepEqual(await i18nextImports(join(packageDir, file)), [], `${ENTRY} ${file}`);
		const main = exports['.'][condition];
		const reached = await i18nextImports(join(packageDir, main));
		assert.ok(
			reached.some((line) => line.includes('react-i18next')),
			`${NAME} ${main} still imports react-i18next`,
		);
		console.log(
			`[binding-pack] ${ENTRY} ${file}: no i18next import; ${main} imports react-i18next`,
		);
	}
}

/** Required peers of the packed packages (the React package's range first), and optional ones. */
function peers() {
	const required = new Map();
	const optional = new Set();
	for (const manifest of manifests) {
		for (const [name, range] of Object.entries(manifest.peerDependencies ?? {})) {
			if (manifest.peerDependenciesMeta?.[name]?.optional) {
				optional.add(name);
			} else if (!required.has(name)) {
				required.set(name, range);
			}
		}
		for (const name of Object.keys(manifest.optionalDependencies ?? {})) {
			optional.add(name);
		}
	}
	return { required, optional: [...optional].filter((name) => !required.has(name)) };
}

const renderCheck = (format) => {
	const load =
		format === 'ESM'
			? `const api = await import('${ENTRY}');
const { renderToString } = await import('react-dom/server');
const { createElement } = await import('react');
const { translationsEn } = await import('${NAME}/i18n');`
			: `const api = require('${ENTRY}');
const { renderToString } = require('react-dom/server');
const { createElement } = require('react');
const { translationsEn } = require('${NAME}/i18n');`;
	return `const assert = ${format === 'ESM' ? "(await import('node:assert/strict')).default" : "require('node:assert/strict')"};
${load}
const loading = translationsEn['pptx.viewer.loading'];
const alone = renderToString(createElement(api.PowerPointViewer, { content: null }));
assert.ok(alone.includes(loading), '${ENTRY} shows its built-in English text outside a provider (${format})');
const translate = (key) => (key === 'pptx.viewer.loading' ? 'Loading (host)' : undefined);
const provided = renderToString(createElement(api.PptxI18nProvider, { locale: 'de', translate }, createElement(api.PowerPointViewer, { content: null })));
assert.ok(provided.includes('Loading (host)'), '${ENTRY} translates through PptxI18nProvider (${format})');
console.log('${ENTRY}: renders without i18next, outside and inside a provider (${format})');
`;
};

/**
 * Install the packed packages without i18next and react-i18next, then render and bundle the
 * entry. `--legacy-peer-deps` stops npm from installing required peers by itself, so the other
 * required peers are installed explicitly at the ranges the packages declare.
 */
export async function checkConsumerWithoutI18next({ directory, tarballs, npm, run }) {
	const { required, optional } = peers();
	for (const name of I18N_PEERS) {
		// Still required in 4.x, so the main entry's hosts keep getting them from npm.
		assert.ok(required.has(name), `${NAME} keeps ${name} as a required peer`);
	}
	mkdirSync(directory);
	writeFileSync(
		join(directory, 'package.json'),
		JSON.stringify({ name: 'pptx-react-no-i18next-consumer', private: true, type: 'module' }),
	);
	const installed = [...required]
		.filter(([name]) => !I18N_PEERS.includes(name))
		.map(([name, range]) => `${name}@${range}`);
	npm(
		[
			'install',
			'--ignore-scripts',
			'--omit=optional',
			'--legacy-peer-deps',
			'--no-audit',
			'--no-fund',
			'--package-lock=false',
			...installed,
			...tarballs,
		],
		directory,
	);
	for (const name of I18N_PEERS) {
		assert.ok(!existsSync(join(directory, 'node_modules', name)), `${name} is not installed`);
	}
	console.log(`[binding-pack] consumer without ${I18N_PEERS.join(' or ')}: ${installed.join(' ')}`);

	writeFileSync(join(directory, 'render.mjs'), renderCheck('ESM'));
	run(process.execPath, ['render.mjs'], directory);
	writeFileSync(join(directory, 'render.cjs'), renderCheck('CommonJS'));
	run(process.execPath, ['render.cjs'], directory);
	writeFileSync(
		join(directory, 'main.mjs'),
		`import assert from 'node:assert/strict';
await assert.rejects(import('${NAME}'), (error) => error.code === 'ERR_MODULE_NOT_FOUND' && error.message.includes("'react-i18next'"));
console.log('${NAME}: the main entry needs react-i18next (ESM)');
`,
	);
	run(process.execPath, ['main.mjs'], directory);

	const bundle = async (specifier, file) => {
		const entry = join(directory, file);
		writeFileSync(
			entry,
			`import { createElement } from 'react';
import { createRoot } from 'react-dom/client';
import { PowerPointViewer, PptxI18nProvider } from '${specifier}';
const translate = (key) => key;
createRoot(document.getElementById('root')).render(
  createElement(PptxI18nProvider, { locale: 'en', translate }, createElement(PowerPointViewer, { content: null })),
);
`,
		);
		return bundleBrowserApp({ entry, external: optional });
	};
	assert.deepEqual(
		await bundle(ENTRY, 'app-no-i18next.js'),
		[],
		`an app importing ${ENTRY} builds`,
	);
	const errors = await bundle(NAME, 'app-main.js');
	assert.ok(errors.length > 0, `an app importing ${NAME} fails without react-i18next`);
	assert.deepEqual(
		[...new Set(errors)],
		['Could not resolve "react-i18next"'],
		`an app importing ${NAME} fails only on react-i18next`,
	);
	console.log(
		`[binding-pack] browser bundle: ${ENTRY} builds without i18next; ${NAME} fails: ${errors[0]}`,
	);
}
