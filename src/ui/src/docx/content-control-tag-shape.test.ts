// @vitest-environment jsdom
import { readFile } from 'node:fs/promises';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { describe, expect, it } from 'vitest';
import { loadDocx } from 'ooxml-core/docx';
import { contentControlTagMetrics } from 'ooxml-core/docx/layout';
import { contentControlPlugin } from 'ooxml-core/docx/ui';
import { contentControlViewPlugin } from './content-control-view';
import { modelToDoc } from './model-adapter';

async function tags(name: string, edit?: (view: EditorView) => void): Promise<HTMLElement[]> {
	const { model } = await loadDocx(
		await readFile(
			new URL(
				'../../../core/docx/layout/fixtures/word-layout-reports/' + name + '.docx',
				import.meta.url,
			),
		),
	);
	const host = document.body.appendChild(document.createElement('div'));
	const view = new EditorView(host, {
		state: EditorState.create({
			doc: modelToDoc(model),
			plugins: [contentControlPlugin(), contentControlViewPlugin()],
		}),
	});
	try {
		edit?.(view);
		return [...view.dom.querySelectorAll<HTMLElement>('.dve-sdt-tag')].map(
			(tag) => tag.cloneNode(true) as HTMLElement,
		);
	} finally {
		view.destroy();
		host.remove();
	}
}

const points = (value: string) => {
	expect(value).toMatch(/^-?[\d.]+pt$/);
	return Number.parseFloat(value);
};

/** How far a tag's margin box reaches above and below the body baseline, in points. */
function lineExtent(tag: HTMLElement): { above: number; below: number } {
	const { height, lineHeight, verticalAlign, marginTop, marginBottom } = tag.style;
	// A one-line box puts its baseline half the line height plus half of Tahoma's
	// ascent less its descent (1 and 0.207 em) below its top.
	const baseline = points(lineHeight) / 2 + ((1 - 0.207) * contentControlTagMetrics.fontSizePt) / 2;
	return {
		above: points(verticalAlign) + baseline + points(marginTop),
		below: points(height) - baseline - points(verticalAlign) + points(marginBottom),
	};
}

const ends = (tag: HTMLElement) =>
	[...tag.querySelectorAll('svg.dve-sdt-tag-end')].map((svg) => ({
		side: svg.getAttribute('data-side'),
		viewBox: svg.getAttribute('viewBox'),
		fill: svg.querySelector('.dve-sdt-tag-fill')?.getAttribute('d'),
		line: svg.querySelector('.dve-sdt-tag-line')?.getAttribute('d'),
		markers: svg.querySelectorAll('.dve-sdt-tag-marker').length,
	}));

describe('content control tag shape in the editing view', () => {
	it("draws Word's arc-sided ends at each side of an inline tag", async () => {
		const [start, end] = await tags('measured-inline-content-control-tags');
		expect(ends(start!)).toEqual([
			{
				side: 'left',
				viewBox: '0 -0.54 2.28 14.52',
				fill: expect.stringMatching(/^M 2\.28 13\.44 C -0\.\d+ [\d.]+ -0\.\d+ [\d.]+ 2\.28 0 Z$/),
				line: expect.stringMatching(/^M 2\.28 0 C -0\.\d+ /),
				markers: 0,
			},
			{
				side: 'right',
				viewBox: '0 -0.54 2.28 14.52',
				fill: expect.stringMatching(/^M 0 0 L 2\.28 0 C -0\.\d+ .* L 0 13\.44 Z$/),
				line: expect.stringMatching(/^M 0 0 L 2\.28 0 C .* L 0 13\.44$/),
				markers: 0,
			},
		]);
		expect(ends(end!).map(({ fill }) => fill?.slice(0, 12))).toEqual([
			'M 2.28 0 L 0',
			'M 0 13.44 C ',
		]);
		expect(start!.style).toMatchObject({
			height: '14.52pt',
			lineHeight: '12.256pt',
			verticalAlign: '3pt',
			minInlineSize: '18.5pt',
			paddingInline: '6.72pt 6.78pt',
		});
		expect(end!.style).toMatchObject({ minInlineSize: '9.84pt', marginInlineEnd: '1.66pt' });
	});

	it('marks paragraph, row and cell tags on their outer end', async () => {
		const all = await tags('block-content-control-tags');
		const block = all.filter((tag) => tag.dataset.tagTitle === 'Block');
		expect(block.map((tag) => ends(tag).map(({ markers }) => markers))).toEqual([
			[2, 0],
			[0, 2],
		]);
		expect(block.map((tag) => tag.style.paddingInline)).toEqual(['9pt 4.5pt', '5.04pt 8.46pt']);
	});

	it("reserves Print Layout's tagged line ascent and descent on every tag", async () => {
		const all = [
			...(await tags('measured-inline-content-control-tags')),
			...(await tags('block-content-control-tags')),
		];
		expect(all.length).toBeGreaterThan(6);
		const { lineAscentPt, lineDescentPt } = contentControlTagMetrics;
		for (const tag of all) {
			const { above, below } = lineExtent(tag);
			expect(above).toBeCloseTo(lineAscentPt, 2);
			expect(below).toBeCloseTo(lineDescentPt, 2);
		}
	});

	it('keeps an exact line height once a paragraph is set to exact spacing', async () => {
		const exact = setParagraphs({ lineSpacingRule: 'exact', lineSpacingTwips: 240 });
		const all = [
			...(await tags('measured-inline-content-control-tags', exact)),
			...(await tags('block-content-control-tags', exact)),
		];
		expect(all.length).toBeGreaterThan(6);
		for (const tag of all) {
			// The margin box collapses onto the baseline, inside the line's own height.
			const { above, below } = lineExtent(tag);
			expect(above).toBeCloseTo(0, 2);
			expect(below).toBeCloseTo(0, 2);
		}
	});

	it('mirrors the tags of a right-to-left paragraph', async () => {
		const rtl = setParagraphs({ direction: 'rtl' });
		const all = [
			...(await tags('measured-inline-content-control-tags', rtl)),
			...(await tags('block-content-control-tags', rtl)),
		];
		expect(all.length).toBeGreaterThan(6);
		for (const tag of all) expect(tag.dataset.tagDirection).toBe('rtl');
		const ltr = await tags('measured-inline-content-control-tags');
		for (const tag of ltr) expect(tag.dataset.tagDirection).toBeUndefined();
	});
});

/** Sets attributes on every paragraph after the view is drawn, as the Paragraph dialog does. */
function setParagraphs(attrs: Record<string, unknown>): (view: EditorView) => void {
	return (view) => {
		let tr = view.state.tr;
		view.state.doc.descendants((node, pos) => {
			if (node.type.name !== 'paragraph') return;
			for (const [name, value] of Object.entries(attrs)) tr = tr.setNodeAttribute(pos, name, value);
		});
		view.dispatch(tr);
	};
}
