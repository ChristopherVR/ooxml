import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { expect, it } from 'vitest';
import { loadDocx, twipsToPixels } from '../index';
import { parseXml } from '../../xml/index';
import {
	createPageStories,
	cssFontStack,
	layoutDocumentModel,
	DEFAULT_FONT_SIZE_PT,
} from './index';
import { createFakeMeasurer } from './measure';
import { paragraphBorderBands } from './adapt-paragraph-box';
import { at } from './test-helpers';

const measurer = createFakeMeasurer();
const load = async (name: string) =>
	loadDocx(await readFile(new URL(`./fixtures/word-layout-reports/${name}.docx`, import.meta.url)));

async function layout(name: string, textMeasurer = measurer) {
	const { model } = await load(name);
	const { pages } = layoutDocumentModel(model, textMeasurer);
	const page = at(pages, 0);
	const blocks = at(page.columns, 0).blocks;
	const paragraphs = blocks.filter((block) => block.kind === 'paragraph');
	const lines = paragraphs.flatMap((block) => block.lines);
	const stories = createPageStories(model, pages, textMeasurer);
	return { model, page, pages, paragraphs, lines, blocks, stories };
}

it('paints legacy group geometry in physical coordinates while keeping the outline and label size (#329)', async () => {
	const { page } = await layout('k4-groupe-formes');
	const svg = at(page.floats, 0).chart?.svg;
	expect(svg).toBeDefined();
	const root = parseXml(svg ?? '').documentElement;
	const paths = Array.from(root.getElementsByTagName('path'));
	expect(paths.map((path) => Number(path.getAttribute('stroke-width')))).toEqual([
		0.33, 0.33, 0.33,
	]);
	expect(at(paths, 1).getAttribute('d')).toMatch(/^M 0 0 L 21[0-9]/);
	expect(root.textContent).toBe('ORDONNANCE');
	expect(root.getElementsByTagName('text')[0]?.getAttribute('font-size')).toBe('16');
});

it.each(['l01-times-12pt', 'l02-arial-12pt'])(
	'flows %s to the same first-page break as Word with no styles part (#330)',
	async (name) => {
		const { pages, paragraphs, lines, model } = await layout(name);
		expect(pages).toHaveLength(7);
		expect(paragraphs).toHaveLength(43);
		const pageParagraphs = pages.map((page) =>
			page.columns.flatMap((column) => column.blocks.filter((block) => block.kind === 'paragraph')),
		);
		expect(pageParagraphs.map((blocks) => blocks.length)).toEqual([43, 43, 43, 43, 43, 43, 42]);
		expect(
			pageParagraphs.flatMap((blocks) =>
				blocks.map((block) => [block.blockId, block.lines.map((line) => line.sourceStart)]),
			),
		).toEqual(Array.from({ length: 300 }, (_, index) => [`p${index}`, [0]]));
		expect(lines[0]?.heightPx).toBeCloseTo(16 / 0.75, 0);
		// Word switches to true single spacing when the package supplies a style catalog
		// with no line spacing. Keep both branches exercised with identical text.
		const single = { ...model, paragraphStyles: { docDefaults: {}, styles: {}, warnings: [] } };
		const result = layoutDocumentModel(single, measurer);
		expect(result.pages).toHaveLength(6);
		expect(result.pages[0]?.columns[0]?.blocks).toHaveLength(50);
	},
);

it.each(['k3-pied-de-page-haut', 'k3-pied-de-page-haut-premiere'])(
	'reserves room for the actual footer on each page of %s (#331)',
	async (name) => {
		const { pages, stories, page } = await layout(name);
		expect(pages).toHaveLength(3);
		const footer = at(stories(0), 0);
		expect(footer.kind).toBe('footer');
		expect(footer.y).toBeCloseTo(page.heightPx - page.marginBottomPx);
		expect(page.marginBottomPx).toBeGreaterThan(1135 / 15);
		for (const column of page.columns)
			for (const block of column.blocks)
				expect(page.marginTopPx + block.yPx + block.heightPx).toBeLessThanOrEqual(footer.y);
		const second = at(pages, 1);
		if (name.endsWith('premiere')) expect(second.marginBottomPx).toBeCloseTo(1135 / 15);
		else expect(second.marginBottomPx).toBeCloseTo(page.marginBottomPx);
	},
);

it('lays out indents and one continuous box around the seven footer paragraphs (#332)', async () => {
	const { stories } = await layout('k3-encadre-pied');
	const boxes = at(stories(0), 0).blocks.filter((box) => box.kind === 'paragraph');
	expect(boxes).toHaveLength(7);
	const first = at(boxes, 0);
	expect(first.frame?.leftPx).toBeGreaterThan(350);
	expect(first.frame?.borders?.top).toBeDefined();
	expect(first.frame?.heightPx).toBeCloseTo(at(boxes, 1).yPx - first.yPx);
	for (const box of boxes.slice(1, -1)) {
		expect(box.frame?.borders?.top).toBeUndefined();
		expect(box.frame?.borders?.bottom).toBeUndefined();
		expect(box.frame?.borders?.left).toBeDefined();
		expect(box.frame?.heightPx).toBeGreaterThan(box.heightPx);
	}
	expect(at(boxes, 6).frame?.borders?.bottom).toBeDefined();
	const simple = await layout('k2-header-footer-border');
	const [header, footer] = simple.stories(0);
	const paragraphFrame = (box: (typeof boxes)[number] | undefined) => box?.frame?.borders;
	expect(
		paragraphFrame(header?.blocks.find((box) => box.kind === 'paragraph'))?.bottom,
	).toBeDefined();
	expect(paragraphFrame(footer?.blocks.find((box) => box.kind === 'paragraph'))?.top).toBeDefined();
});

it('places a header table float from the cell, matching the body control (#333)', async () => {
	const header = await layout('k2-anchor-in-cell-header');
	const body = await layout('k2-anchor-in-cell');
	const picture = at(at(header.stories(0), 0).floats, 0);
	expect(picture.inCell).toBe(true);
	expect(picture.xPx * 0.75).toBeCloseTo(420.75, 0);
	expect(at(body.page.floats, 0).xPx * 0.75).toBeCloseTo(421, 0);
});

it('produces triple and thin/thick border bands in their native order (#334)', async () => {
	const { paragraphs } = await layout('k2-border-compound');
	const bands = paragraphs.flatMap((paragraph) => {
		const line = paragraph.frame?.borders?.top;
		return line?.compound ? [paragraphBorderBands(line, false).map((band) => band.widthPx)] : [];
	});
	expect(bands).toEqual([
		[2 / 3, 2 / 3, 2 / 3],
		[2 / 3, 1],
		[1, 2 / 3],
		[2 / 3, 1, 2 / 3],
	]);
});

it('does not move an inline picture above its line at three-quarter spacing (#335)', async () => {
	const { lines } = await layout('k2-inline-image-line-multiple');
	const pictures = lines.flatMap((line) => line.fragments.filter((fragment) => fragment.object));
	expect(pictures).toHaveLength(2);
	expect(pictures.map((fragment) => fragment.topPx)).toEqual([0, 0]);
	expect(at(pictures, 0).object?.heightPx).toEqual(at(pictures, 1).object?.heightPx);
});

it('right-aligns the preceding line in the usable space beside the picture (#336)', async () => {
	const { page, lines } = await layout('k3-alignement-droite-habillage');
	const picture = at(page.floats, 0);
	const line = lines.find((candidate) =>
		candidate.fragments.some((fragment) => fragment.text === 'Sous'),
	);
	const first = at(line?.fragments, 0);
	const last = line?.fragments.at(-1);
	expect(first.xPx).toBeGreaterThan(100);
	expect(page.marginLeftPx + (last?.xPx ?? 0) + (last?.widthPx ?? 0)).toBeLessThanOrEqual(
		picture.xPx,
	);
});

it('tries the font table alternate followed by a matching generic family (#337)', async () => {
	const { lines } = await layout('k3-police-repli-roman');
	const text = lines
		.flatMap((line) => line.fragments)
		.find((fragment) => fragment.fontFamily === 'Fakeserif Std');
	const stack = cssFontStack(text?.fontFamily, text?.fontFallback);
	expect(stack).toBe('"Fakeserif Std", "Cambria", "Caladea", serif');
});

it('uses explicit fixed pitch and keeps proportional modern fonts proportional (#337)', async () => {
	const cases = [
		{ name: 'Montserrat', family: 'modern', pitch: 'variable', expected: 'sans-serif' },
		{ name: 'Modern Default', family: 'modern', pitch: 'default', expected: 'sans-serif' },
		{ name: 'Modern Unspecified', family: 'modern', expected: 'sans-serif' },
		{ name: 'Modern Fixed', family: 'modern', pitch: 'fixed', expected: 'monospace' },
		{ name: 'Swiss Fixed', family: 'swiss', pitch: 'fixed', expected: 'monospace' },
		{ name: 'Unclassified Fixed', pitch: 'fixed', expected: 'monospace' },
		{ name: 'Roman Variable', family: 'roman', pitch: 'variable', expected: 'serif' },
	];
	const ns = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
	const zip = new JSZip();
	const fonts = cases.map(({ name, family, pitch }) => {
		const fontFamily = family ? `<w:family w:val="${family}"/>` : '';
		const fontPitch = pitch ? `<w:pitch w:val="${pitch}"/>` : '';
		return `<w:font w:name="${name}">${fontFamily}${fontPitch}</w:font>`;
	});
	const paragraphs = cases.map(
		({ name }) =>
			`<w:p><w:r><w:rPr><w:rFonts w:ascii="${name}" w:hAnsi="${name}"/></w:rPr><w:t>Wi</w:t></w:r></w:p>`,
	);
	zip.file('word/fontTable.xml', `<w:fonts ${ns}>${fonts.join('')}</w:fonts>`);
	zip.file(
		'word/document.xml',
		`<w:document ${ns}><w:body>${paragraphs.join('')}</w:body></w:document>`,
	);
	const { model } = await loadDocx(await zip.generateAsync({ type: 'uint8array' }));
	const { pages } = layoutDocumentModel(model, measurer);
	const stacks = at(at(pages, 0).columns, 0).blocks.flatMap((block) =>
		block.kind === 'paragraph'
			? block.lines.flatMap((line) =>
					line.fragments.map(({ fontFamily, fontFallback }) =>
						cssFontStack(fontFamily, fontFallback),
					),
				)
			: [],
	);
	expect(stacks).toEqual(cases.map(({ name, expected }) => `"${name}", ${expected}`));
});

it('uses the same height for omitted size and explicit 10 pt (#338)', async () => {
	const { lines } = await layout('k4-taille-par-defaut');
	expect(DEFAULT_FONT_SIZE_PT).toBe(10);
	expect(at(lines, 0).heightPx).toBeCloseTo(at(lines, 1).heightPx);
	expect(at(lines, 0).heightPx).toBeLessThan(at(lines, 2).heightPx);
});

it('paints percentage shading behind paragraphs and table cells and preserves it on save (#339)', async () => {
	const { paragraphs, blocks } = await layout('k4-trame-motif');
	expect(paragraphs.find((paragraph) => paragraph.frame)?.frame?.shading).toBe('#e5e5e5');
	const table = blocks.find((block) => block.kind === 'table');
	expect(table?.rows[0]?.geometry?.map((cell) => cell.shading)).toEqual([
		'#e5e5e5',
		'#bfbfbf',
		'#7f7f7f',
	]);
	const loaded = await load('k4-trame-motif');
	const first = at(loaded.model.blocks, 0);
	if (first.type !== 'paragraph') throw new Error('Expected an introduction');
	const model = {
		...loaded.model,
		blocks: [
			{ ...first, runs: [...first.runs, { text: ' Edited.' }] },
			...loaded.model.blocks.slice(1),
		],
	};
	const zip = await JSZip.loadAsync(await loaded.save(model));
	const xml = await zip.file('word/document.xml')?.async('string');
	expect(xml).toContain(' Edited.');
	for (const percentage of [10, 25, 50]) expect(xml).toContain(`w:val="pct${percentage}"`);
});

it('carries yellow highlight and small caps into the painted runs after tabs and in cells (#340)', async () => {
	const { lines, blocks } = await layout('k4-surlignage-petites-capitales');
	const flat = lines.flatMap((line) => line.fragments);
	expect(flat.some((fragment) => fragment.text && fragment.background === '#ffff00')).toBe(true);
	expect(flat.some((fragment) => fragment.text && fragment.smallCaps)).toBe(true);
	const table = blocks.find((block) => block.kind === 'table');
	const cellFragments = table?.rows.flatMap((row) =>
		row.cells.flatMap((cell) => cell.flatMap((p) => p.lines.flatMap((line) => line.fragments))),
	);
	expect(
		cellFragments?.some((fragment) => fragment.text && fragment.background === '#ffff00'),
	).toBe(true);
	expect(cellFragments?.some((fragment) => fragment.text && fragment.smallCaps)).toBe(true);
});

it('keeps an uncached merge field in the source without displaying a placeholder (#341)', async () => {
	const { lines, model } = await layout('k4-champ-simple-vide');
	expect(
		lines
			.flatMap((line) => line.fragments)
			.map((fragment) => fragment.text)
			.join(''),
	).not.toContain('[Field]');
	expect(
		model.blocks.some(
			(block) =>
				block.type === 'paragraph' &&
				block.runs.some((run) => run.fieldCode?.includes('MERGEFIELD')),
		),
	).toBe(true);
});

it('uses the document interval and explicit tab stops in the footer (#342)', async () => {
	const { stories, model } = await layout('k4-tabulations-pied');
	expect(model.defaultTabStopTwips).toBe(284);
	const footer = at(stories(0), 0);
	const lines = footer.blocks.flatMap((block) => (block.kind === 'paragraph' ? block.lines : []));
	const defaultTab = at(lines, 0).fragments.find((fragment) => fragment.text.startsWith('par'));
	expect(defaultTab?.xPx).toBeGreaterThan(250);
	const explicit = lines
		.flatMap((line) => line.fragments)
		.find((fragment) => fragment.text === 'ici');
	expect(explicit?.xPx).toBe(2835 / 15);
});

it('measures empty paragraphs at the paragraph mark font size (#343)', async () => {
	const { lines } = await layout('k4-paragraphe-vide-marque');
	for (const line of lines.slice(1, 7)) expect(line.heightPx).toBeCloseTo(2 * (4 / 3) * 1.15);
	for (const line of lines.slice(9, 15)) expect(line.heightPx).toBeCloseTo(8 * (4 / 3) * 1.15);
});

it('extends a horizontal paragraph border to the hanging first line (#344)', async () => {
	const { paragraphs, model } = await layout('k3-trait-paragraphe-vide');
	const [normal, hanging] = paragraphs.filter((paragraph) => paragraph.frame?.borders?.bottom);
	const source = model.blocks.find(
		(block) =>
			block.type === 'paragraph' && block.borders?.bottom && block.hangingTwips !== undefined,
	);
	expect(source?.type).toBe('paragraph');
	const indent = source?.type === 'paragraph' ? source.hangingTwips : undefined;
	expect(indent).toBeDefined();
	expect(hanging?.frame?.leftPx).toBeCloseTo(
		(normal?.frame?.leftPx ?? 0) - (indent ? twipsToPixels(indent) : 0),
	);
});

it('omits a hidden table cell end marker from row height, while keeping explicit content (#330)', async () => {
	const { paragraphs } = await layout('cell-end-marker');
	// Word COM PDF baselines: 84.02, 114.02, 144.02, 159.02, 189.02, 234.02, 279.05 pt.
	const offsets = paragraphs.map((paragraph) => (paragraph.yPx - at(paragraphs, 0).yPx) * 0.75);
	for (const [index, expected] of [0, 30, 60, 75, 105, 150, 195].entries())
		expect(at(offsets, index)).toBeCloseTo(expected, 1);

	const loaded = await load('cell-end-marker');
	const first = at(loaded.model.blocks, 0);
	if (first.type !== 'paragraph') throw new Error('Expected a body paragraph');
	const model = {
		...loaded.model,
		blocks: [{ ...first, runs: [{ text: 'Edited' }] }, ...loaded.model.blocks.slice(1)],
	};
	const saved = await loadDocx(await loaded.save(model));
	const { pages } = layoutDocumentModel(saved.model, measurer);
	const edited = at(at(pages, 0).columns, 0).blocks.filter((block) => block.kind === 'paragraph');
	expect(at(edited, 0).lines[0]?.fragments[0]?.text).toBe('Edited');
	expect(edited.map((paragraph) => (paragraph.yPx - at(edited, 0).yPx) * 0.75)).toEqual(offsets);
});

it('suppresses contextual spacing between implicit Normal paragraphs (#347)', async () => {
	const implicit = await layout('issue-347-contextual-spacing-implicit-normal');
	const explicit = await layout('issue-347-contextual-spacing-explicit-normal-control');
	expect(explicit.pages).toHaveLength(1);
	expect(explicit.paragraphs).toHaveLength(32);
	for (const [index, paragraph] of explicit.paragraphs.slice(1).entries()) {
		const previous = at(explicit.paragraphs, index);
		expect(paragraph.yPx).toBeCloseTo(previous.yPx + previous.heightPx);
	}
	expect(implicit.pages).toHaveLength(1);
	expect(implicit.paragraphs.map(({ yPx, heightPx }) => [yPx, heightPx])).toEqual(
		explicit.paragraphs.map(({ yPx, heightPx }) => [yPx, heightPx]),
	);
});
