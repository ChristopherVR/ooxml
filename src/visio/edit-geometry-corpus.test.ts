import { createHash } from 'node:crypto';
import { readdirSync, readFileSync } from 'node:fs';
import { join, relative, resolve } from 'node:path';
import JSZip from 'jszip';
import { DOMParser } from '@xmldom/xmldom';
import { describe, expect, it } from 'vitest';
import { editVsdx, type VisioEdit, type EditVsdxResult } from './edit.js';
import { VisioPackageError } from './package-common.js';
import { parseVsdx } from './parser.js';

// Opt-in public corpus only. Download/hash provenance belongs to the caller;
// this test never fetches files or depends on another repository's checkout.
const directory = process.env['VISIO_EDIT_CORPUS_DIR'];
function drawings(root: string): string[] {
	return readdirSync(root, { withFileTypes: true })
		.flatMap((entry) => {
			const path = join(root, entry.name);
			return entry.isDirectory() ? drawings(path) : /\.vsdx$/i.test(entry.name) ? [path] : [];
		})
		.sort();
}
const refusals = new Set([
	'UNSUPPORTED_GEOMETRY_EDIT',
	'EDIT_PROTECTED_CELL',
	'EDIT_UNKNOWN_DEPENDENCY',
	'EDIT_DYNAMIC_DEPENDENCY',
	'EDIT_UNSUPPORTED_DEPENDENCY',
	'EDIT_UNSUPPORTED_FORMULA',
	'EDIT_UNSUPPORTED_PACKAGE_DEPENDENCY',
	'EDIT_AMBIGUOUS_CELL',
	'INVALID_SHAPE_ID',
	'EDIT_FORMULA_CYCLE',
	'EDIT_FORMULA_UNIT',
	'EDIT_FORMULA_ERROR',
	'EDIT_FORMULA_VALUE',
	'EDIT_REFERENCED_DELETE',
	'UNSUPPORTED_EDIT_PACKAGE',
	'UNSUPPORTED_XML_EDIT',
]);
async function preserved(original: Uint8Array, result: EditVsdxResult) {
	const before = await JSZip.loadAsync(original),
		after = await JSZip.loadAsync(result.bytes);
	const names = (zip: JSZip) =>
		Object.values(zip.files)
			.filter((entry) => !entry.dir)
			.map((entry) => entry.name)
			.sort();
	expect(names(after)).toEqual(names(before));
	expect(result.changedParts.length).toBeGreaterThan(0);
	for (const [path, entry] of Object.entries(before.files)) {
		if (entry.dir || result.changedParts.includes(path)) continue;
		expect(await after.file(path)!.async('uint8array'), path).toEqual(
			await entry.async('uint8array'),
		);
	}
	return parseVsdx(result.bytes);
}

describe.skipIf(!directory)('public VSDX geometry edit corpus', () => {
	it('reopens admitted edits, preserves untouched payloads and rejects unsupported edits atomically', async () => {
		const root = resolve(directory!);
		const files = drawings(root);
		expect(files.length).toBeGreaterThan(0);
		const results: {
			file: string;
			sha256: string;
			create: string;
			move: string;
			resize: string;
			delete: string;
			candidate?: { pageId: string; shapeId: string; name: string; width: number; height: number };
			reasons?: Partial<Record<'create' | 'move' | 'resize' | 'delete', string>>;
		}[] = [];
		const counts: Record<string, number> = {};
		for (const path of files) {
			const original = new Uint8Array(readFileSync(path)),
				snapshot = original.slice();
			let document: Awaited<ReturnType<typeof parseVsdx>>;
			try {
				document = await parseVsdx(original);
			} catch (error) {
				expect(error, path).toBeInstanceOf(VisioPackageError);
				const code = (error as VisioPackageError).code;
				// The pinned public corpus includes malformed ZIP and missing-part cases.
				expect(['MISSING_PART', 'INVALID_ZIP'], path).toContain(code);
				await expect(
					editVsdx(original, [
						{
							type: 'create-rectangle',
							pageId: '0',
							shapeId: '1',
							x: 2,
							y: 2,
							width: 1,
							height: 1,
						},
					]),
				).rejects.toMatchObject({ code });
				expect(original).toEqual(snapshot);
				counts[`admission:${code}`] = (counts[`admission:${code}`] ?? 0) + 1;
				results.push({
					file: relative(root, path),
					sha256: createHash('sha256').update(original).digest('hex'),
					create: `admission:${code}`,
					move: 'package-not-admitted',
					resize: 'package-not-admitted',
					delete: 'package-not-admitted',
				});
				continue;
			}
			const page = document.pages[0];
			expect(page, path).toBeDefined();
			const ids = new Set<string>();
			const collect = (shapes: NonNullable<typeof page>['shapes']) => {
				for (const shape of shapes) {
					ids.add(shape.id);
					collect(shape.children);
				}
			};
			collect(page!.shapes);
			let id = 1;
			while (ids.has(String(id))) id++;
			const record: (typeof results)[number] = {
				file: relative(root, path),
				sha256: createHash('sha256').update(original).digest('hex'),
				create: '',
				move: 'no-local-2d-candidate',
				resize: 'no-local-2d-candidate',
				delete: 'no-local-2d-candidate',
				reasons: {},
			};
			const attempt = async (name: 'create' | 'move' | 'resize' | 'delete', edits: VisioEdit[]) => {
				let result: EditVsdxResult;
				try {
					result = await editVsdx(original, edits);
				} catch (error) {
					expect(error, `${record.file}: ${name}`).toBeInstanceOf(VisioPackageError);
					const structured = error as VisioPackageError;
					expect(
						refusals.has(structured.code),
						`${record.file}: unexpected ${structured.code}: ${structured.message}`,
					).toBe(true);
					expect(structured.message.length).toBeGreaterThan(0);
					expect(original).toEqual(snapshot);
					record[name] = structured.code;
					record.reasons![name] = structured.message;
					counts[`${name}:${structured.code}`] = (counts[`${name}:${structured.code}`] ?? 0) + 1;
					return;
				}
				expect(original).toEqual(snapshot);
				const reopened = await preserved(snapshot, result);
				const changed = reopened.pages.find((candidate) => candidate.id === page!.id)!;
				if (name === 'create') {
					const created = changed.shapes.find((shape) => shape.id === String(id))!;
					expect(created).toBeDefined();
					expect([created.width, created.height, created.text.plainText]).toEqual([
						1.5,
						0.75,
						'Corpus rectangle',
					]);
				} else if (name === 'delete') {
					expect(changed.shapes.some((shape) => shape.id === edits[0]!.shapeId)).toBe(false);
				} else if (name === 'resize') {
					const target = changed.shapes.find((shape) => shape.id === edits[0]!.shapeId)!;
					const resize = edits[0] as Extract<VisioEdit, { type: 'resize-shape' }>;
					expect([target.width, target.height]).toEqual([resize.width, resize.height]);
				} else {
					const zip = await JSZip.loadAsync(result.bytes);
					const parse = (xml: string) => new DOMParser().parseFromString(xml, 'text/xml');
					const pages = parse(await zip.file('visio/pages/pages.xml')!.async('string'));
					const pageNode = Array.from(pages.getElementsByTagName('Page')).find(
						(node) => node.getAttribute('ID') === page!.id,
					)!;
					const relationId = pageNode.getElementsByTagName('Rel')[0]!.getAttribute('r:id');
					const relations = parse(
						await zip.file('visio/pages/_rels/pages.xml.rels')!.async('string'),
					);
					const relationship = Array.from(relations.getElementsByTagName('Relationship')).find(
						(node) => node.getAttribute('Id') === relationId,
					)!;
					const part = relationship.getAttribute('Target')!;
					const pagePath = part.startsWith('/') ? part.slice(1) : `visio/pages/${part}`;
					const contents = parse(await zip.file(pagePath)!.async('string'));
					const target = Array.from(contents.getElementsByTagName('Shape')).find(
						(node) => node.getAttribute('ID') === edits[0]!.shapeId,
					)!;
					const cells = Array.from(target.getElementsByTagName('Cell'));
					const numeric = (name: string) =>
						Number(cells.find((node) => node.getAttribute('N') === name)!.getAttribute('V'));
					expect([numeric('PinX'), numeric('PinY')]).toEqual([2, 3]);
				}
				record[name] = 'success';
				counts[`${name}:success`] = (counts[`${name}:success`] ?? 0) + 1;
			};
			await attempt('create', [
				{
					type: 'create-rectangle',
					pageId: page!.id,
					shapeId: String(id),
					x: 2,
					y: 2,
					width: 1.5,
					height: 0.75,
					text: 'Corpus rectangle',
				},
			]);
			const target = page!.shapes.find(
				(shape) =>
					shape.kind === 'shape' &&
					!shape.masterId &&
					shape.children.length === 0 &&
					shape.width > 0 &&
					shape.height > 0,
			);
			if (target) {
				record.candidate = {
					pageId: page!.id,
					shapeId: target.id,
					name: target.name,
					width: target.width,
					height: target.height,
				};
				await attempt('move', [
					{ type: 'move-shape', pageId: page!.id, shapeId: target.id, x: 2, y: 3 },
				]);
				await attempt('resize', [
					{
						type: 'resize-shape',
						pageId: page!.id,
						shapeId: target.id,
						width: target.width * 1.1,
						height: target.height * 1.1,
					},
				]);
				await attempt('delete', [{ type: 'delete-shape', pageId: page!.id, shapeId: target.id }]);
			} else
				counts['existing:no-local-2d-candidate'] =
					(counts['existing:no-local-2d-candidate'] ?? 0) + 1;
			results.push(record);
		}
		console.info(
			'VSDX geometry corpus results',
			JSON.stringify({ files: files.length, counts, results }),
		);
	}, 120_000);
});
