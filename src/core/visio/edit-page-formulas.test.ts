import { describe, expect, it } from 'vitest';
import { readFile, writeFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { editVsdx } from './edit.js';
import { evaluateVisioFormula } from './formula.js';
import { cell, fixture, shape } from './test-fixtures.js';
import { VisioPackage } from './package.js';
import { indexedPart, related, visioXml } from './parts.js';
import { attribute, children } from './sheet.js';
import evidence from './__fixtures__/page-formulas-native.json';

const user = (name: string, value: number, formula: string) =>
	`<Row N="${name}">${cell('Value', value, formula)}</Row>`;
const fields = (number: number, extra = '') =>
	`<Section N="User">${user('Number', number, 'PAGENUMBER()')}${user('Count', 2, 'PAGECOUNT()')}${user('Combined', number + 2, 'User.Number+User.Count')}${extra}</Section>`;
const source = (extra = '') =>
	fixture({
		pages: [
			{ id: '10', contents: `<Shapes>${shape('1', fields(1, extra))}</Shapes>` },
			{
				id: '20',
				attributes: 'Background="1"',
				contents: `<Shapes>${shape('1', fields(0))}</Shapes>`,
			},
			{ id: '30', contents: `<Shapes>${shape('1', fields(2))}</Shapes>` },
		],
	});
it('matches the recorded native page-function results without a local Visio installation', () => {
	for (const page of evidence.recalculated) {
		const context = { pageNumber: page.number, pageCount: page.count };
		const resolve = () => ({ value: 0, unit: 'scalar' as const });
		expect(evaluateVisioFormula('PAGENUMBER()+PAGECOUNT()', resolve, context).value).toBe(
			page.combined,
		);
		const accepted = evidence.accepted.find((item) => item.id === page.id)!;
		expect(accepted).toEqual(page);
	}
});
async function values(bytes: Uint8Array, path: string) {
	const xml = await (await JSZip.loadAsync(bytes)).file(path)!.async('string');
	return [
		...xml.matchAll(/<Row[^>]*N=["'](Number|Count|Combined)["'][^>]*>\s*<Cell[^>]*V=["']([^"']+)/g),
	].map((match) => [match[1], Number(match[2])]);
}
it('requires explicit context and zero arguments for page functions', () => {
	const resolve = () => ({ value: 0, unit: 'scalar' as const });
	expect(
		evaluateVisioFormula('PAGENUMBER()+PAGECOUNT()', resolve, { pageNumber: 0, pageCount: 2 }),
	).toEqual({ value: 2, unit: 'scalar' });
	expect(() => evaluateVisioFormula('PAGECOUNT()', resolve)).toThrow(/context/);
	expect(() => evaluateVisioFormula('PAGENUMBER(1)', resolve, { pageNumber: 1 })).toThrow(
		/argument count/,
	);
});
it('refreshes foreground page numbers and transitive caches while preserving F', async () => {
	const saved = await editVsdx(await source(), [{ type: 'reorder-page', pageId: '30', index: 0 }]);
	expect(await values(saved.bytes, 'visio/pages/page1.xml')).toEqual([
		['Number', 2],
		['Count', 2],
		['Combined', 4],
	]);
	expect(await values(saved.bytes, 'visio/pages/page2.xml')).toEqual([
		['Number', 0],
		['Count', 2],
		['Combined', 2],
	]);
	expect(await values(saved.bytes, 'visio/pages/page3.xml')).toEqual([
		['Number', 1],
		['Count', 2],
		['Combined', 3],
	]);
	const xml = await (
		await JSZip.loadAsync(saved.bytes)
	)
		.file('visio/pages/page3.xml')!
		.async('string');
	expect(xml).toContain('F="PAGENUMBER()"');
	expect(saved.changedParts).not.toContain('visio/pages/page2.xml');
});
it('refreshes foreground count and dependent cells on foreground and background pages', async () => {
	const saved = await editVsdx(await source(), [
		{ type: 'insert-page', pageId: '40', afterPageId: '10', name: 'New' },
	]);
	expect(await values(saved.bytes, 'visio/pages/page1.xml')).toEqual([
		['Number', 1],
		['Count', 3],
		['Combined', 4],
	]);
	expect(await values(saved.bytes, 'visio/pages/page2.xml')).toEqual([
		['Number', 0],
		['Count', 3],
		['Combined', 3],
	]);
	expect(await values(saved.bytes, 'visio/pages/page3.xml')).toEqual([
		['Number', 3],
		['Count', 3],
		['Combined', 6],
	]);
});
it('refreshes PageSheet caches without putting PageSheet markup in page contents', async () => {
	const original = await fixture({
		pages: [
			{ id: '0', pageCells: fields(1), contents: '<Shapes/>' },
			{ id: '1', contents: '<Shapes/>' },
		],
	});
	const saved = await editVsdx(original, [
		{ type: 'insert-page', pageId: '2', afterPageId: '0', name: 'New' },
	]);
	const zip = await JSZip.loadAsync(saved.bytes);
	expect(await values(saved.bytes, 'visio/pages/pages.xml')).toContainEqual(['Count', 3]);
	expect(await zip.file('visio/pages/page1.xml')!.async('string')).not.toContain('PageSheet');
	expect(await zip.file('visio/pages/page3.xml')!.async('string')).not.toContain('PageSheet');
});
it.each([user('Unsupported', 1, 'FONT(User.Count)'), user('Cycle', 1, 'User.Count+User.Cycle')])(
	'refuses unprovable affected closures atomically',
	async (extra) => {
		const original = await source(extra),
			snapshot = new Uint8Array(original);
		await expect(
			editVsdx(original, [{ type: 'insert-page', pageId: '40', afterPageId: '10', name: 'New' }]),
		).rejects.toThrow();
		expect(original).toEqual(snapshot);
	},
);
it('refuses inherited page-dependent caches', async () => {
	const original = await fixture({ masters: [{ id: '1', shapes: shape('1', fields(1)) }] });
	await expect(
		editVsdx(original, [{ type: 'insert-page', pageId: '40', afterPageId: '0', name: 'New' }]),
	).rejects.toThrow(/Inherited page-dependent/);
});
it('keeps identical page-order transactions byte-identical', async () => {
	const original = await source();
	expect(
		(await editVsdx(original, [{ type: 'reorder-page', pageId: '10', index: 0 }])).bytes,
	).toEqual(original);
});

const native = process.env.VISIO_NATIVE_PAGE_FORMULAS_DIR;
describe.skipIf(!native)('native numeric page formulas', () => {
	it('matches freshly recalculated native foreground and background caches', async () => {
		const original = await readFile(resolve(native!, 'page-formulas.vsdx'));
		const saved = await editVsdx(original, [{ type: 'reorder-page', pageId: '5', index: 1 }]);
		await writeFile(resolve(native!, 'core-reordered.vsdx'), saved.bytes);
		const inserted = await editVsdx(original, [
			{ type: 'insert-page', pageId: '40', afterPageId: '0', name: 'New' },
		]);
		await writeFile(resolve(native!, 'core-inserted.vsdx'), inserted.bytes);
		const referenceBytes = await readFile(resolve(native!, 'recalculated-native.vsdx'));
		const reference = await JSZip.loadAsync(referenceBytes);
		const output = await JSZip.loadAsync(saved.bytes);
		const paths = async (bytes: Uint8Array) => {
			const pkg = await VisioPackage.open(bytes);
			const document = (await related(pkg, '', 'document'))!;
			const part = (await related(pkg, document, 'pages'))!;
			const result = new Map<string, string>();
			for (const page of children(await visioXml(pkg, part, 'Pages'), 'Page'))
				result.set(attribute(page, 'ID')!, await indexedPart(pkg, part, page, 'page'));
			return result;
		};
		const referencePaths = await paths(referenceBytes),
			outputPaths = await paths(saved.bytes);
		for (const [id, path] of outputPaths) {
			const caches = (xml: string) =>
				[
					...xml.matchAll(
						/<Cell N=["']Value["'][^>]*V=["']([^"']+)["'][^>]*F=["'](?:PAGENUMBER\(\)|PAGECOUNT\(\)|User.Number\+User.Count)["']/g,
					),
				].map((m) => Number(m[1]));
			const expected = caches(await reference.file(referencePaths.get(id)!)!.async('string'));
			expect(expected).toHaveLength(3);
			expect(caches(await output.file(path)!.async('string'))).toEqual(expected);
		}
	});
});
