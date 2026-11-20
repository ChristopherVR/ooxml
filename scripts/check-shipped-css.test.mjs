import assert from 'node:assert/strict';
import { mkdirSync, mkdtempSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { test } from 'node:test';

import { cssProblems, exportedCss, packedFiles, packagesUnder } from './check-shipped-css.mjs';

// The tail of pptx-vue-viewer@4.80.0's stylesheet: a comment whose opening `/*` (and the text up
// to a `}` inside it) was lost, leaving the rest of the comment as raw CSS.
const BROKEN = `@media (max-width: 767px) {
select[data-v-6d24293a] {
		min-height: 44px;
}
}
\`) is UNLAYERED CSS, and an unlayered rule always
 * wins over a Tailwind utility class (which lives inside \`@layer utilities\`)
 * specificity beats the \`:where()\` baseline on both axes at once, which is
 * why it must set both properties, not just the one sub-panels forget.
 */
@media (max-width: 767px) {
.pptx-vue-inspector[data-v-e5d19d7e] select {
		min-height: 44px;
}
}
`;

test('a comment that lost its opening delimiter fails both minifiers', async () => {
	const problems = await cssProblems(BROKEN, 'pptx-vue-viewer.css');
	assert.ok(
		problems.some((problem) => problem.startsWith('esbuild: Unexpected "`"')),
		problems,
	);
	assert.ok(
		problems.some((problem) => problem.startsWith('lightningcss: ')),
		problems,
	);
});

test('valid CSS with comments, layers and nesting passes', async () => {
	const css = `@layer theme, utilities;
/* a comment with { braces } and \`backticks\` */
@layer utilities { .a { color: red; &:hover { color: blue } } }
@media (max-width: 767px) { .b[data-v-1] { min-height: 44px } }
`;
	assert.deepEqual(await cssProblems(css, 'ok.css'), []);
});

test('an unterminated block is reported', async () => {
	const problems = await cssProblems('.a { color: red;', 'open.css');
	assert.ok(problems.length > 0);
});

test('a prefix selects the publishable packages under it, with their publish directory', () => {
	const pptx = packagesUnder(['viewers/pptx']);
	const names = pptx.map((pkg) => pkg.npm);
	assert.ok(names.includes('pptx-vue-viewer'));
	assert.ok(pptx.every((pkg) => pkg.dir.startsWith('viewers/pptx/')));
	const angular = pptx.find((pkg) => pkg.npm === 'pptx-angular-viewer');
	assert.match(angular.packDir.replaceAll('\\', '/'), /packages\/angular\/dist$/);
	assert.deepEqual(
		packagesUnder(['src/ui']).map((pkg) => pkg.npm),
		['ooxml-ui'],
	);
});

test('Tailwind directives in a stylesheet published for Tailwind builds are not problems', async () => {
	const css = `@theme inline { --color-primary: red; }
@source '../viewer';
.a { color: red }
`;
	assert.deepEqual(await cssProblems(css, 'theme.css'), []);
});

test('the stylesheets a manifest exports, at any condition depth', () => {
	const manifest = {
		exports: {
			'.': { import: './dist/index.js' },
			'./styles': './dist/viewer.css',
			'./styles.css': './dist/viewer.css',
			'./theme.css': { default: './src/theme.css' },
		},
	};
	assert.deepEqual(exportedCss(manifest), ['dist/viewer.css', 'src/theme.css']);
});

test('packed files follow the files field and never include node_modules', () => {
	const dir = mkdtempSync(join(tmpdir(), 'shipped-css-'));
	try {
		for (const path of [
			'dist/a.css',
			'dist/sub/b.css',
			'src/theme.css',
			'src/other.css',
			'dist/node_modules/x.css',
		]) {
			mkdirSync(join(dir, path, '..'), { recursive: true });
			writeFileSync(join(dir, path), '');
		}
		const files = packedFiles(dir, { files: ['dist/', 'src/theme.css', 'missing'] }).sort();
		assert.deepEqual(files, ['dist/a.css', 'dist/sub/b.css', 'src/theme.css']);
		assert.ok(packedFiles(dir, {}).includes('src/other.css'));
	} finally {
		rmSync(dir, { recursive: true, force: true });
	}
});
