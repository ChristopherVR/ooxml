import { test, expect, type Locator } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { join } from 'node:path';
import { parseVsdx, type VisioEdit } from 'ooxml-core/visio';
import { composeAffine, IDENTITY_AFFINE, type AffineMatrix } from 'ooxml-core/geometry';
interface Tree {
	id: string;
	transform: number[];
	cells: Record<string, { value: number }>;
	children: Tree[];
}
type Host = HTMLElement & {
	applyEdits(edits: readonly VisioEdit[]): Promise<void>;
	undo(): Promise<void>;
	redo(): Promise<void>;
	exportVsdx(): { bytes: Uint8Array };
};
async function compareDom(viewer: Locator, node: Tree, ratio: number): Promise<void> {
	const transforms = await viewer.locator(`[data-shape-id="${node.id}"]`).evaluate((element) => {
		const result: number[][] = [];
		for (let ancestor: Element | null = element; ancestor; ancestor = ancestor.parentElement) {
			if (!ancestor.hasAttribute('data-shape-id')) continue;
			const source = ancestor.getAttribute('transform')!;
			if (!/^matrix\([^()]+\)$/.test(source))
				throw new Error('Expected an authored affine matrix.');
			result.unshift(
				source
					.slice(7, -1)
					.trim()
					.split(/[\s,]+/)
					.map(Number),
			);
		}
		return result;
	});
	let authored: AffineMatrix = IDENTITY_AFFINE;
	for (const transform of transforms) {
		expect(transform).toHaveLength(6);
		authored = composeAffine(authored, [
			transform[0]!,
			transform[1]!,
			transform[2]!,
			transform[3]!,
			transform[4]!,
			transform[5]!,
		]);
	}
	for (let i = 0; i < 6; i++)
		expect(authored[i]).toBeCloseTo(node.transform[i]! * (i >= 4 ? ratio : 1), 12);
	const matrix = await viewer.locator(`[data-shape-id="${node.id}"]`).evaluate((element) => {
		const shape = element as SVGGraphicsElement,
			paper = shape.closest('svg')!;
		const matrix = paper.getScreenCTM()!.inverse().multiply(shape.getScreenCTM()!);
		const height =
			Number(paper.getAttribute('data-page-height')) ||
			(paper as SVGSVGElement).viewBox.baseVal.height;
		return [matrix.a, -matrix.b, matrix.c, -matrix.d, matrix.e, height - matrix.f];
	});
	for (let i = 0; i < 6; i++)
		// Chromium exposes float-rounded screen matrices; authored geometry is checked above.
		expect(matrix[i]).toBeCloseTo(node.transform[i]! * (i >= 4 ? ratio : 1), 5);
	for (const child of node.children) await compareDom(viewer, child, ratio);
}
for (const variable of ['VISIO_NATIVE_GROUP_ROTATE_DIR', 'VISIO_NATIVE_GROUP_ROTATE_NESTED_DIR'])
	for (const framework of ['vanilla', 'react', 'vue', 'angular', 'svelte', 'solid'])
		test(`${framework}: native group API rotation/history/reload (${variable})`, async ({
			page,
		}) => {
			const directory = process.env[variable];
			test.skip(!directory, `Set ${variable} to a native group rotation capture.`);
			const source = await readFile(join(directory!, 'source.vsdx'));
			const evidence = JSON.parse(await readFile(join(directory!, 'evidence.json'), 'utf8')) as {
				source: Tree;
				rotated: Tree;
				pageScale: number;
				drawingScale: number;
			};
			const ratio = evidence.pageScale / evidence.drawingScale;
			await page.goto(framework === 'vanilla' ? '/demo/?sample=1' : `/demo-${framework}/?sample=1`);
			await page.locator('#file').setInputFiles({
				name: 'group.vsdx',
				mimeType: 'application/vnd.ms-visio.drawing',
				buffer: source,
			});
			await expect(page.locator('#file-name')).toHaveText('group.vsdx');
			const viewer = page.locator('visio-viewer');
			await compareDom(viewer, evidence.source, ratio);
			await viewer.evaluate(
				(element, { id, angle }) =>
					(element as Host).applyEdits([{ type: 'rotate-shape', pageId: '0', shapeId: id, angle }]),
				{ id: evidence.source.id, angle: evidence.rotated.cells.Angle!.value },
			);
			await compareDom(viewer, evidence.rotated, ratio);
			const exported = async () =>
				Buffer.from(
					await viewer.evaluate((element) => Array.from((element as Host).exportVsdx().bytes)),
				);
			const saved = await exported();
			const originalModel = await parseVsdx(source),
				model = await parseVsdx(saved);
			expect(model.pages[0]!.shapes[0]!.children).toEqual(
				originalModel.pages[0]!.shapes[0]!.children,
			);
			const compare = (
				shape: (typeof model.pages)[number]['shapes'][number],
				node: Tree,
				parent: AffineMatrix = IDENTITY_AFFINE,
			) => {
				const world = composeAffine(parent, shape.transform);
				for (let i = 0; i < 6; i++)
					expect(world[i]).toBeCloseTo(node.transform[i]! * (i >= 4 ? ratio : 1), 12);
				for (const child of shape.children)
					compare(
						child,
						node.children.find((node) => node.id === child.id)!,
						world,
					);
			};
			compare(model.pages[0]!.shapes[0]!, evidence.rotated);
			await viewer.evaluate((element) => (element as Host).undo());
			expect(await exported()).toEqual(source);
			await compareDom(viewer, evidence.source, ratio);
			await viewer.evaluate((element) => (element as Host).redo());
			expect(await exported()).toEqual(saved);
			await compareDom(viewer, evidence.rotated, ratio);
			await page.locator('#file').setInputFiles({
				name: 'group-copy.vsdx',
				mimeType: 'application/vnd.ms-visio.drawing',
				buffer: saved,
			});
			await expect(page.locator('#file-name')).toHaveText('group-copy.vsdx');
			await compareDom(viewer, evidence.rotated, ratio);
		});
