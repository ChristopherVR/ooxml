#!/usr/bin/env node
// `vite build` compiles every Vue SFC's `<style>` block into dist/pptx-vue-viewer.css. The build
// then runs `@tailwindcss/cli` to compile the SAME output path from src/styles/pptx-vue-viewer.css,
// which knows nothing of the SFC styles and overwrites them (issue #196). So "save" sets Vite's
// stylesheet aside before the Tailwind pass and "merge" appends it, whole, after the Tailwind
// output.
//
// The SFC stylesheet is appended byte for byte, never split or filtered: an earlier version split it
// into rules by counting braces, which a `{ ... }` inside a comment threw off, so the comment lost
// its opening `/*` and shipped as raw text that bundlers could not minify.
//
// Usage:
//   node vue-css-merge.mjs save  <cssPath>   (run right after `vite build`)
//   node vue-css-merge.mjs merge <cssPath>   (run right after the Tailwind CLI pass)
import { existsSync, readFileSync, renameSync, unlinkSync, writeFileSync } from 'node:fs';

/** The Tailwind output followed by the SFC stylesheet. Throws if the SFC styles are missing. */
export function mergeScopedCss(tailwindCss, sfcCss) {
	// Guard against a repeat of issue #196: if Vite ever emits no scoped styles (a Vue or Vite
	// upgrade, a config change), every `<style scoped>` rule silently vanishes from the package.
	if (!sfcCss.includes('data-v-')) {
		throw new Error(
			'[vue-css-merge] the SFC stylesheet has no [data-v-*] scoped rules - the SFC ' +
				'<style scoped> blocks were lost. Refusing to ship (see issue #196).',
		);
	}
	return `${tailwindCss}\n${sfcCss}`;
}

function main() {
	const [, , mode, cssPath] = process.argv;
	const snapshotPath = `${cssPath}.vue-sfc-snapshot`;

	if (mode === 'save') {
		renameSync(cssPath, snapshotPath);
		console.log(`[vue-css-merge] set aside the SFC stylesheet`);
	} else if (mode === 'merge') {
		if (!existsSync(snapshotPath)) {
			throw new Error(
				`[vue-css-merge] missing snapshot at ${snapshotPath}; did the "save" step run first?`,
			);
		}
		const sfcCss = readFileSync(snapshotPath, 'utf8');
		writeFileSync(cssPath, mergeScopedCss(readFileSync(cssPath, 'utf8'), sfcCss));
		unlinkSync(snapshotPath);
		console.log(`[vue-css-merge] appended ${sfcCss.length} bytes of SFC CSS`);
	} else {
		throw new Error('usage: node vue-css-merge.mjs <save|merge> <cssPath>');
	}
}

if (process.argv[1]?.endsWith('vue-css-merge.mjs')) {
	main();
}
