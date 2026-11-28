import { readFile } from 'node:fs/promises';
import JSZip from 'jszip';
import { expect, it, vi } from 'vitest';
import { loadDocx, twipsToPixels } from '../index';
import { parseXml } from '../../xml/index';
import {
	createPageStories,
	adaptDocumentModel,
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
		return line?.compound ? [paragraphBorderBands(line).map((band) => band.widthPx)] : [];
	});
	// Word's bands at `w:sz` 4: 0.48 pt (0.64 px) wide, the thin band 0.72 pt (0.96 px), and the
	// wide band of a thin-thick style first (#388).
	const w = 0.64;
	const expected = [
		[w, w],
		[w, w, w],
		[w, 0.96],
		[0.96, w],
		[0.96, w, 0.96],
	];
	expect(bands).toHaveLength(expected.length);
	bands.forEach((band, index) => {
		expect(band).toHaveLength(expected[index]?.length ?? -1);
		band.forEach((width, bandIndex) =>
			expect(width).toBeCloseTo(expected[index]?.[bandIndex] ?? NaN, 6),
		);
	});
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

it('drops contextual spacing only from the flagged paragraph (#359)', async () => {
	const rows = await layout('issue-359-contextual-spacing-rows');
	const word = JSON.parse(
		await readFile(
			new URL(
				'./fixtures/word-layout-reports/issue-359-contextual-spacing-rows.json',
				import.meta.url,
			),
			'utf8',
		),
	) as { gapsPt: Record<string, number> };
	const wordContextualGapsPt = word.gapsPt;
	const byText = new Map(
		rows.paragraphs.map((paragraph) => [
			paragraph.lines.flatMap((line) => line.fragments.map((part) => part.text ?? '')).join(''),
			paragraph,
		]),
	);
	const gaps = Object.fromEntries(
		Object.keys(wordContextualGapsPt).map((id) => {
			const upper = byText.get(`${id}U`);
			const lower = byText.get(`${id}L`);
			if (!upper || !lower) throw new Error(`missing pair ${id}`);
			return [id, Math.round((lower.yPx - upper.yPx - upper.heightPx) * 0.75 * 100) / 100];
		}),
	);
	expect(gaps).toEqual(wordContextualGapsPt);

	const oneSided = await layout('issue-359-contextual-spacing-one-sided');
	const control = await layout('issue-359-contextual-spacing-both-control');
	expect(oneSided.pages).toHaveLength(2);
	expect(oneSided.paragraphs).toHaveLength(33);
	expect(control.pages).toHaveLength(1);
	expect(control.paragraphs).toHaveLength(40);
	for (const [index, paragraph] of control.paragraphs.slice(1).entries()) {
		const previous = at(control.paragraphs, index);
		expect(paragraph.yPx).toBeCloseTo(previous.yPx + previous.heightPx);
	}
	// Each flag-free "Lorem ipsum" paragraph keeps its 10 pt after; the flagged one drops its own.
	const lorem = at(oneSided.paragraphs, 0);
	const consectetur = at(oneSided.paragraphs, 1);
	const next = at(oneSided.paragraphs, 2);
	expect((consectetur.yPx - lorem.yPx - lorem.heightPx) * 0.75).toBeCloseTo(10);
	expect(next.yPx).toBeCloseTo(consectetur.yPx + consectetur.heightPx);
});

/** Each line's text, page and baseline in points from the top of its page. */
function lineBaselines(pages: ReturnType<typeof layoutDocumentModel>['pages']) {
	return pages.flatMap((page) =>
		page.columns.flatMap((column) =>
			column.blocks.flatMap((block) =>
				block.kind === 'paragraph'
					? block.lines.map((line) => ({
							text: line.fragments.map((part) => part.text ?? '').join(''),
							page: page.index,
							pt: (page.marginTopPx + block.yPx + line.yPx + (line.baselinePx ?? 0)) * 0.75,
						}))
					: [],
			),
		),
	);
}

const wordLineSpacing = async () =>
	JSON.parse(
		await readFile(
			new URL('./fixtures/word-layout-reports/issue-360-line-spacing.json', import.meta.url),
			'utf8',
		),
	) as {
		lineSpacingRules: { baselinesPt: Record<string, number> };
		tripleSpacingPageFit: { pages: number; linesPerPage: number[]; firstBaselinePt: number };
		lastLine: { linesOnPageOne: number; baselinesPt: number[] };
		lastLineControl: { linesOnPageOne: number };
	};

// Calibri's single line at 12 pt is 14.64 pt in Word; the fake measurer's default is 13.8 pt.
const calibriLike = createFakeMeasurer({ lineHeightFactor: 14.64 / 12 });

it('puts the extra space of multiple line spacing below the text (#360)', async () => {
	const word = await wordLineSpacing();
	const { pages } = await layout('issue-360-multiple-line-spacing-last-line', calibriLike);
	expect(pages).toHaveLength(1);
	const lines = lineBaselines(pages);
	expect(lines).toHaveLength(word.lastLine.linesOnPageOne);
	const steps = (values: number[]) =>
		values.slice(1).map((value, index) => value - at(values, index));
	const ours = steps(lines.map((line) => line.pt));
	const theirs = steps(word.lastLine.baselinesPt);
	// Line 4, the first 1.15 line, steps like a single line; Word's PDF rounds to 0.12 pt.
	expect(at(theirs, 2)).toBeCloseTo(14.64, 2);
	ours.forEach((step, index) => expect(Math.abs(step - at(theirs, index))).toBeLessThan(0.13));
	const control = await layout('issue-360-multiple-line-spacing-last-line-control', calibriLike);
	expect(control.pages).toHaveLength(1);
	expect(lineBaselines(control.pages)).toHaveLength(word.lastLineControl.linesOnPageOne);
});

it('keeps Word baselines for auto and at-least line spacing (#360)', async () => {
	const word = (await wordLineSpacing()).lineSpacingRules.baselinesPt;
	const { pages } = await layout('issue-360-line-spacing-rules');
	const ours = new Map(lineBaselines(pages).map((line) => [line.text, line.pt]));
	const labels = Object.keys(word);
	// Steps into and within single, multiple and at-least groups. Word puts an exact line's baseline
	// 80% down the line and shrinks both extents below single spacing; those groups (m080, exact20
	// and the lines after them) are known gaps, outside #360.
	const known = new Set([
		'm0800',
		'm0801',
		'm0802',
		'm080s',
		'exact200',
		'exact201',
		'exact202',
		'exact20s',
		'least200',
	]);
	for (const [index, label] of labels.slice(1).entries()) {
		if (known.has(label)) continue;
		const previous = at(labels, index);
		const step = (ours.get(label) ?? NaN) - (ours.get(previous) ?? NaN);
		expect(
			Math.abs(step - (at(Object.values(word), index + 1) - at(Object.values(word), index))),
		).toBeLessThan(0.13);
	}
});

it('lets the extra space of the last line on a page pass the bottom margin (#360)', async () => {
	const word = (await wordLineSpacing()).tripleSpacingPageFit;
	const { pages } = await layout('issue-360-triple-spacing-page-fit');
	expect(pages).toHaveLength(word.pages);
	const lines = lineBaselines(pages);
	expect(pages.map((page) => lines.filter((line) => line.page === page.index).length)).toEqual(
		word.linesPerPage,
	);
	// The first line of each page sits one ascent below the top margin, as at single spacing.
	for (const page of pages) {
		const first = lines.find((line) => line.page === page.index);
		expect(Math.abs((first?.pt ?? NaN) - word.firstBaselinePt)).toBeLessThan(0.13);
	}
});

it('lays out zero character scale at natural width and preserves the source value (#348)', async () => {
	const loaded = await load('issue-348-character-scale-zero');
	const control = await load('issue-348-character-scale-zero-control');
	const result = (model: typeof loaded.model) => {
		const { pages } = layoutDocumentModel(model, measurer);
		return pages.map((page) =>
			page.columns.flatMap((column) =>
				column.blocks.flatMap((block) =>
					block.kind === 'paragraph'
						? [
								block.lines.map((line) =>
									line.fragments.map(({ text, widthPx }) => [text, widthPx]),
								),
							]
						: [],
				),
			),
		);
	};
	const expected = result(control.model);
	expect(expected).toHaveLength(2);
	expect(expected.flat().filter((paragraph) => paragraph.length > 1)).toHaveLength(12);
	const actual = result(loaded.model);
	expect(actual.map((page) => page.map((paragraph) => paragraph.length))).toEqual(
		expected.map((page) => page.map((paragraph) => paragraph.length)),
	);
	expect(actual).toEqual(expected);
	const first = at(loaded.model.blocks, 0);
	if (first.type !== 'paragraph') throw new Error('Expected a paragraph');
	const edited = {
		...loaded.model,
		blocks: [
			{ ...first, runs: [...first.runs, { text: ' Edit' }] },
			...loaded.model.blocks.slice(1),
		],
	};
	const xml = await JSZip.loadAsync(await loaded.save(edited));
	expect(await xml.file('word/document.xml')?.async('string')).toContain('<w:w w:val="0"');
});

it('applies TableGrid paragraph spacing only inside its table (#353)', async () => {
	const tableMeasurer = createFakeMeasurer({ charWidthFactor: 0.48 });
	const control = await layout('issue-353-table-direct-spacing-control', tableMeasurer);
	const styled = await layout('issue-353-table-style-spacing', tableMeasurer);
	for (const sample of [control, styled]) {
		const [table, after] = at(adaptDocumentModel(sample.model).sections, 0).blocks;
		if (table?.kind !== 'table') throw new Error('Expected a table');
		const paragraphs = at(at(table.rows, 0).cells, 0).paragraphs;
		expect(paragraphs).toHaveLength(24);
		for (const paragraph of paragraphs)
			expect(paragraph).toMatchObject({
				spacingAfterTwips: 0,
				lineSpacingTwips: 240,
				lineSpacingRule: 'auto',
			});
		expect(after).toMatchObject({
			kind: 'paragraph',
			spacingAfterTwips: 160,
			lineSpacingTwips: 259,
		});
		expect(sample.pages).toHaveLength(1);
	}
	const tableHeight = (sample: typeof control) => {
		const table = sample.blocks.find((block) => block.kind === 'table');
		return at(table?.rows, 0).heightPx;
	};
	expect(tableHeight(styled)).toBeCloseTo(tableHeight(control));
});

it('lets table justification override only left justification from the default style (#353)', async () => {
	const loaded = await load('issue-353-table-style-precedence');
	const { model } = loaded;
	const cellFormatting = (source: typeof model) => {
		const [body, table] = at(adaptDocumentModel(source).sections, 0).blocks;
		expect(body).toMatchObject({ align: 'left', spacingAfterTwips: 240 });
		if (table?.kind !== 'table') throw new Error('Expected a table');
		return at(at(table.rows, 0).cells, 0).paragraphs.map((p) => [p.align, p.spacingAfterTwips]);
	};
	expect(cellFormatting(model)).toEqual([
		['center', 240],
		['center', 240],
		['left', 240],
		['right', 360],
	]);
	const edited = await loadDocx(
		await loaded.save({ ...model, overrideTableStyleFontSizeAndJustification: true }),
	);
	expect(edited.model.overrideTableStyleFontSizeAndJustification).toBe(true);
	expect(cellFormatting(edited.model)).toEqual([
		['left', 240],
		['left', 240],
		['left', 240],
		['right', 360],
	]);
});

it.each([
	['issue-358-justified-line-break-donotexpandshiftreturn', true],
	['issue-358-justified-line-break-control', false],
] as const)('lays out soft breaks in %s (#358)', async (name, noExpansion) => {
	const { page, lines, paragraphs } = await layout(name);
	const columnWidth = at(page.columns, 0).widthPx;
	const lineEnd = (line: (typeof lines)[number]) => {
		const last = at(line.fragments, line.fragments.length - 1);
		return last.xPx + last.widthPx;
	};
	// This first line wraps at the available width instead of ending in a soft break.
	expect(lineEnd(at(at(paragraphs, 0).lines, 0))).toBeCloseTo(columnWidth);
	const text = 'And for the other party, represented by:';
	const line = at(at(paragraphs, 1).lines, 0);
	expect(line.fragments.map((f) => f.text).join('')).toBe(text);
	expect(at(line.fragments, 0).xPx).toBe(0);
	const first = at(line.fragments, 0);
	const font = { family: first.fontFamily ?? 'Calibri', sizePx: (first.fontSizePt ?? 11) / 0.75 };
	expect(lineEnd(line)).toBeCloseTo(noExpansion ? measurer.widthOf(text, font) : columnWidth);
	if (noExpansion) {
		for (const space of line.fragments.filter((f) => f.text === ' '))
			expect(space.widthPx).toBeCloseTo(measurer.widthOf(' ', font));
	}
});

it.each([
	['distribute', true],
	['thaiDistribute', false],
	['mediumKashida', false],
] as const)(
	'expands soft-break lines under doNotExpandShiftReturn only for %s (#358)',
	async (justification, expands) => {
		const { model } = await load('issue-358-justified-line-break-donotexpandshiftreturn');
		const blocks = model.blocks.map((block) =>
			block.type === 'paragraph' ? { ...block, justification } : block,
		);
		const { pages } = layoutDocumentModel({ ...model, blocks }, measurer);
		const column = at(at(pages, 0).columns, 0);
		const paragraph = at(column.blocks, 1);
		if (paragraph.kind !== 'paragraph') throw new Error('Expected a paragraph');
		const fragments = at(paragraph.lines, 0).fragments;
		const last = at(fragments, fragments.length - 1);
		const end = last.xPx + last.widthPx;
		if (expands) expect(end).toBeCloseTo(column.widthPx);
		// Natural width of the 40-character line under the test measurer.
		else expect(end).toBeCloseTo(330.88);
	},
);

it('preserves the soft line break setting through a document edit (#358)', async () => {
	const loaded = await load('issue-358-justified-line-break-donotexpandshiftreturn');
	const first = at(loaded.model.blocks, 0);
	if (first.type !== 'paragraph') throw new Error('Expected a body paragraph');
	const saved = await loadDocx(
		await loaded.save({
			...loaded.model,
			blocks: [{ ...first, runs: [{ text: 'Edited' }] }, ...loaded.model.blocks.slice(1)],
		}),
	);
	expect(saved.model.doNotExpandShiftReturn).toBe(true);
	expect(at(saved.model.blocks, 0)).toMatchObject({ runs: [{ text: 'Edited' }] });
});

it.each([
	['issue-354-float-wrap-none-crosses-bottom-margin', true],
	['issue-354-float-wrap-square-crosses-bottom-margin', true],
	['issue-354-float-fits-above-bottom-margin-control', false],
] as const)(
	'keeps the anchor paragraph and picture on page one in %s (#354)',
	async (name, overlapsMargin) => {
		const { page, pages, paragraphs } = await layout(name);
		expect(pages).toHaveLength(1);
		expect(paragraphs.slice(-2).map((block) => block.blockId)).toEqual(['p24', 'p25']);
		expect(page.floats).toHaveLength(1);
		const picture = at(page.floats, 0);
		expect(picture.blockId).toBe('p24');
		expect(picture.yPx).toBeCloseTo(862.47);
		const bottom = picture.yPx + picture.heightPx;
		expect(bottom).toBeLessThan(page.heightPx);
		expect(bottom > page.heightPx - page.marginBottomPx).toBe(overlapsMargin);
	},
);

it.each([
	['issue-354-float-wrap-none-crosses-bottom-margin', 1],
	['issue-354-float-wrap-square-crosses-bottom-margin', 2],
] as const)('paginates %s in compatibility mode 15 (#354)', async (name, pageCount) => {
	const { model } = await load(name);
	const { pages } = layoutDocumentModel({ ...model, compatibilityMode: 15 }, measurer);
	expect(pages).toHaveLength(pageCount);
	expect(at(at(pages, pageCount - 1).floats, 0).blockId).toBe('p24');
});

it('lays out numbered continuous section breaks like the unnumbered markers (#355)', async () => {
	const control = await layout('issue-355-unnumbered-section-break-paragraph');
	const numbered = await layout('issue-355-numbered-section-break-paragraph');
	const paragraphLayout = (sample: typeof control) =>
		sample.pages.map((page) =>
			page.columns.flatMap((column) =>
				column.blocks.flatMap((block) =>
					block.kind === 'paragraph'
						? [
								{
									id: block.blockId,
									y: block.yPx,
									lines: block.lines.map((line) =>
										line.fragments.map((fragment) => fragment.text).join(''),
									),
								},
							]
						: [],
				),
			),
		);
	const expected = paragraphLayout(control);
	expect(expected).toHaveLength(1);
	const firstPage = at(expected, 0);
	expect(firstPage.map(({ id }) => id)).not.toContain('p1');
	expect(firstPage.map(({ id }) => id)).not.toContain('p3');
	expect(firstPage.find(({ id }) => id === 'p2')?.lines[0]).toMatch(/^1\.\s*First heading$/);
	expect(firstPage.find(({ id }) => id === 'p4')?.lines[0]).toMatch(/^1\.1\s*First subheading$/);
	expect(firstPage.find(({ id }) => id === 'p8')?.lines[0]).toMatch(/^1\.2\s*Second subheading$/);
	expect(paragraphLayout(numbered)).toEqual(expected);
});

it('leaves no line for a numbered next-page section break and restarts text on page two (#355)', async () => {
	const { pages, paragraphs } = await layout('issue-355-numbered-section-break-nextpage');
	expect(pages).toHaveLength(2);
	expect(paragraphs.map(({ blockId }) => blockId)).toEqual(['p0']);
	const firstItem = at(at(at(pages, 1).columns, 0).blocks, 0);
	if (firstItem.kind !== 'paragraph') throw new Error('Expected the first numbered item');
	expect(firstItem.blockId).toBe('p2');
	expect(
		at(firstItem.lines, 0)
			.fragments.map(({ text }) => text)
			.join(''),
	).toMatch(/^1\.\s*Item 1/);
});

it.each([
	['issue-356-control-no-mark-size', 18, 738.5],
	['issue-356-control-with-text', 10, 591.3],
	['issue-356-empty-numbered-paragraphs', 10, 591.3],
])(
	'uses %s marker sizing and keeps the end of the list in place (#356)',
	async (name, size, endY) => {
		const { pages, paragraphs } = await layout(name);
		expect(pages).toHaveLength(1);
		const end = at(paragraphs, 13);
		expect(
			at(end.lines, 0)
				.fragments.map(({ text }) => text)
				.join(''),
		).toMatch(/^End of list:/);
		expect(end.yPx).toBeCloseTo(endY, 0);
		for (const [index, paragraph] of paragraphs.slice(1, 13).entries()) {
			expect(paragraph.lines).toHaveLength(1);
			const line = at(paragraph.lines, 0);
			expect(at(line.fragments, 0)).toMatchObject({ text: `${index + 1}.`, fontSizePt: size });
			expect(line.heightPx).toBeCloseTo(size * (4 / 3) * 1.15);
		}
	},
);

it.each([
	[
		'issue-357-date-field-fr-FR',
		[
			'Run language fr-FR: 26 novembre 2026',
			'Run language fr-FR: jeudi 26 novembre 2026',
			'No run language, document default fr-FR: 26 novembre 2026',
		],
	],
	[
		'issue-357-date-field-en-US-control',
		['Run language en-US: 26 November 2026', 'Run language en-US: Thursday 26 November 2026'],
	],
])('localizes complex DATE results in %s (#357)', async (name, expected) => {
	vi.useFakeTimers({ toFake: ['Date'] });
	try {
		vi.setSystemTime(new Date(2026, 10, 26, 12));
		const { paragraphs } = await layout(name);
		expect(
			paragraphs.map((paragraph) =>
				paragraph.lines.flatMap((line) => line.fragments.map(({ text }) => text)).join(''),
			),
		).toEqual(expected);
	} finally {
		vi.useRealTimers();
	}
});
