/**
 * What a screen reader gets from each document, read from Chromium's accessibility tree: what the
 * products expose today is asserted outright, and the structure they do not expose yet runs as an
 * expected failure owned by the package that adds it (`expected-failures.json`).
 */
import { EXCEL, POWERPOINT, WORD } from './support/fixtures';
import { gate } from './support/known';
import { expect, openSurface, test } from './support/suite';
import { axTree, nodes, outline } from './support/tree';

test.use({ viewport: { width: 1440, height: 900 } });

test.describe('Word accessibility tree', () => {
	test.beforeEach(async ({ page }) => {
		await openSurface(page, 'docx');
	});

	test('the picture is named by its alt text and the link by its text', async ({ page }) => {
		const tree = await axTree(page);
		expect(nodes(tree, 'image', WORD.alt)).toHaveLength(1);
		expect(nodes(tree, 'link', WORD.link)).toHaveLength(1);
	});

	test('headings are exposed with their levels', async ({ page }) => {
		await gate('tree/docx/headings', async () => {
			const tree = await axTree(page);
			const headings = nodes(tree, 'heading').map((node) => [node.name, node.properties.level]);
			const listing = outline(tree, ['heading', 'paragraph']);
			expect(headings, `headings found: ${headings.length}\n${listing}`).toEqual(
				expect.arrayContaining([
					[WORD.heading1, 1],
					[WORD.heading2, 2],
				]),
			);
		});
	});

	test('the table header row is exposed as column headers', async ({ page }) => {
		await gate('tree/docx/table-headers', async () => {
			const tree = await axTree(page);
			const headers = nodes(tree, 'columnheader').map((node) => node.name);
			const listing = outline(tree, ['table', 'row', 'columnheader', 'cell']);
			expect(headers, `column headers found: ${headers.length}\n${listing}`).toEqual(
				expect.arrayContaining([...WORD.header]),
			);
		});
	});

	test('the editing area is a named, multi-line text box', async ({ page }) => {
		await gate('tree/docx/editor-name', async () => {
			const tree = await axTree(page);
			const editors = nodes(tree, 'textbox').filter(
				(node) => node.properties.editable === 'richtext',
			);
			expect(
				editors.map((node) => [Boolean(node.name), node.properties.multiline]),
				`rich-text text boxes found: ${editors.length}`,
			).toEqual([[true, true]]);
		});
	});
});

test.describe('Excel accessibility tree', () => {
	test.beforeEach(async ({ page }) => {
		await openSurface(page, 'xlsx');
	});

	test('the active cell is named by its address and value as it moves', async ({ page }) => {
		const focused = async () =>
			(await axTree(page)).filter((node) => node.properties.focused === true).at(-1)?.name;
		await expect.poll(focused).toBe(`A1, ${EXCEL.header[0]}`);
		await page.keyboard.press('ArrowRight');
		await expect.poll(focused).toBe(`B1, ${EXCEL.header[1]}`);
	});

	test('the picture is named by its alt text', async ({ page }) => {
		expect(nodes(await axTree(page), 'image', EXCEL.pictureAlt)).toHaveLength(1);
	});

	test('cells are grid cells in rows', async ({ page }) => {
		await gate('tree/xlsx/gridcells', async () => {
			const tree = await axTree(page);
			const listing = outline(tree, ['grid', 'row', 'gridcell', 'columnheader']);
			const rows = nodes(tree, 'row').length;
			expect(rows, `rows in the grid: ${rows}\n${listing}`).toBeGreaterThan(0);
			expect(nodes(tree, 'gridcell', '100'), listing).toHaveLength(1);
		});
	});
});

test.describe('PowerPoint accessibility tree', () => {
	test.beforeEach(async ({ page }) => {
		await openSurface(page, 'pptx');
	});

	test('the described picture is named and the decorative one is hidden', async ({ page }) => {
		const tree = await axTree(page);
		expect(nodes(tree, 'image', POWERPOINT.alt)).toHaveLength(1);
		expect(tree.filter((node) => /Decorative Picture/u.test(node.name))).toEqual([]);
	});

	test('the slide is named with its number', async ({ page }) => {
		await gate('tree/pptx/slide-name', async () => {
			const slides = (await axTree(page)).filter(
				(node) => node.properties.roledescription === 'slide',
			);
			const names = slides.map((node) => node.name);
			expect(names, `slide names: ${JSON.stringify(names)}`).toEqual([
				expect.stringMatching(/^Slide 1\b/u),
			]);
		});
	});
});
