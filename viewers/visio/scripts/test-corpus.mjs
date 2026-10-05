import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { parseVsdx, VisioPackageError } from 'ooxml-core/visio';
import { assertViewableDocument } from '../dist/scene-validation.js';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const baseline = JSON.parse(await readFile(resolve(root, 'tests/corpus-baseline.json'), 'utf8'));
if (![4, 5].includes(process.argv.length) || (process.argv[4] && process.argv[4] !== '--svg'))
	throw new Error('Usage: node scripts/test-corpus.mjs LIBVISIO_CHECKOUT POI_CHECKOUT [--svg]');
let exportPageSvg, dom;
if (process.argv[4]) {
	const { JSDOM } = await import('jsdom');
	dom = new JSDOM('<!doctype html>');
	globalThis.document = dom.window.document;
	globalThis.XMLSerializer = dom.window.XMLSerializer;
	({ exportPageSvg } = await import('../dist/export-svg.js'));
}
const totals = { accepted: 0, rejected: 0, pages: 0, shapes: 0, paths: 0, svgPages: 0 };
for (const [index, group] of baseline.groups.entries()) {
	const input = resolve(process.argv[index + 2]);
	for (const file of group.files) {
		const label = `${group.id}/${file.path}`;
		const bytes = await readFile(resolve(input, file.path));
		assert.equal(bytes.length, file.bytes, `${label}: byte count differs from pinned source`);
		assert.equal(
			createHash('sha256').update(bytes).digest('hex'),
			file.sha256,
			`${label}: SHA-256 differs`,
		);
		if (file.rejectCode) {
			await assert.rejects(
				() => parseVsdx(bytes),
				(error) => error instanceof VisioPackageError && error.code === file.rejectCode,
				`${label}: malformed input must reject with its structured error`,
			);
			totals.rejected++;
			continue;
		}
		const model = await parseVsdx(bytes);
		assertViewableDocument(model);
		let shapes = 0,
			paths = 0;
		const stack = model.pages.flatMap((page) => page.shapes);
		while (stack.length) {
			const shape = stack.pop();
			shapes++;
			paths += shape.geometry.length;
			stack.push(...shape.children);
		}
		assert.deepEqual(
			{ pages: model.pages.length, shapes, paths },
			{ pages: file.pages, shapes: file.shapes, paths: file.paths },
			`${label}: normalized regression baseline changed; review before updating`,
		);
		// Facts asserted in the pinned upstream libvisio import tests. These selected
		// colors are not a visual oracle for every effect, font or page placement.
		if (group.id === 'libvisio' && file.path.endsWith('/blue-box.vsdx'))
			assert.equal(model.pages[0].shapes[0].style.fill, '#5b9bd5');
		if (group.id === 'libvisio' && file.path.endsWith('/color-boxes.vsdx'))
			assert.deepEqual(
				model.pages[0].shapes.map((shape) => shape.style.fill),
				['#759fcc', '#70ad47', '#fec000', '#41719c', '#ed7d31', '#bdd0e9', '#5b9bd5'],
			);
		if (group.id === 'poi' && file.path.endsWith('/60973.vsdx'))
			for (const id of ['1', '2', '4', '5', '9']) {
				const shape = model.pages[0].shapes.find((candidate) => candidate.id === id);
				assert.ok(shape, `Expected rounded source shape ${id}`);
				assert.equal((shape.geometry[0].path.match(/\bA\b/g) ?? []).length, 4);
			}
		totals.accepted++;
		totals.pages += model.pages.length;
		totals.shapes += shapes;
		totals.paths += paths;
		if (exportPageSvg)
			for (const [pageIndex] of model.pages.entries()) {
				const result = exportPageSvg(model, pageIndex);
				const parsed = new dom.window.DOMParser().parseFromString(result.svg, 'image/svg+xml');
				assert.equal(parsed.querySelector('parsererror'), null, `${label}: valid SVG XML`);
				if (
					group.id === 'poi' &&
					file.path.endsWith('/60973.vsdx') &&
					model.pages[pageIndex].id === '0'
				) {
					assert.equal(
						parsed.querySelectorAll('marker').length,
						8,
						'60973: eight saved code-5 ends',
					);
					assert.equal(
						parsed.querySelectorAll('[marker-start][marker-end]').length,
						4,
						'60973: four two-ended connectors',
					);
				}

				assert.equal(
					parsed.querySelectorAll('[data-shape-id], [tabindex], script, foreignObject, a, style')
						.length,
					0,
					`${label}: no host controls, executable content or external styles`,
				);
				for (const element of parsed.querySelectorAll('*'))
					for (const attribute of element.attributes) {
						assert.ok(!attribute.name.startsWith('on'), `${label}: no event attributes`);
						if (attribute.localName === 'href')
							assert.ok(
								attribute.value.startsWith('#') ||
									/^data:image\/(png|jpeg|gif);base64,/.test(attribute.value),
								`${label}: local resources only`,
							);
					}
				assert.ok(
					parsed.querySelector('metadata')?.textContent.includes('diagnostics'),
					`${label}: diagnostics survive export`,
				);
				totals.svgPages++;
			}
		console.log(`PASS ${label}: ${model.pages.length} pages, ${shapes} shapes, ${paths} paths`);
	}
}
console.log(JSON.stringify(totals));
console.log(
	'Hash-pinned parse/scene regressions passed. Native Visio visual parity remains unverified.',
);
dom?.window.close();
