// @vitest-environment jsdom
import { readFile } from 'node:fs/promises';
import { EditorState } from 'prosemirror-state';
import { EditorView } from 'prosemirror-view';
import { describe, expect, it } from 'vitest';
import { loadDocx } from 'ooxml-core/docx';
import { contentControlPlugin } from 'ooxml-core/docx/ui';
import { contentControlViewPlugin } from './content-control-view';
import { modelToDoc } from './model-adapter';

async function tags(name: string): Promise<HTMLElement[]> {
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
		return [...view.dom.querySelectorAll<HTMLElement>('.dve-sdt-tag')].map(
			(tag) => tag.cloneNode(true) as HTMLElement,
		);
	} finally {
		view.destroy();
		host.remove();
	}
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
});
