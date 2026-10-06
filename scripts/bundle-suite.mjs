/**
 * Bundle `ooxml-ui/suite` (the host-neutral tab model of the suite shell) into one browser module,
 * so the static launcher in `site/` can import it without a build step of its own. The Pages
 * build writes it next to the launcher; `preview-site.mjs` writes it into `site/` for local review.
 */
import { createRequire } from 'node:module';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const ROOT = resolve(dirname(fileURLToPath(import.meta.url)), '..');

/** @param {string} outfile where to write the module */
export async function bundleSuite(outfile) {
	// tsup is a dependency of ooxml-ui, so resolve it from there (the workspace is isolated).
	const ui = createRequire(join(ROOT, 'src', 'ui', 'package.json'));
	const { build } = await import(pathToFileURL(ui.resolve('tsup')).href);
	await build({
		entry: { suite: join(ROOT, 'src', 'ui', 'src', 'suite', 'index.ts') },
		outDir: dirname(outfile),
		format: ['esm'],
		target: 'es2022',
		dts: false,
		clean: false,
		splitting: false,
		silent: true,
		config: false,
	});
}
