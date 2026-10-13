// @vitest-environment jsdom
import { URL as NodeURL } from 'node:url';
import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import { loadDocx } from 'ooxml-core/docx';
import { layoutDocumentModel } from 'ooxml-core/docx/layout';
import { renderPrintLayout } from './print-layout';
const evidence = JSON.parse(
	readFileSync(
		new NodeURL(
			'../../../core/docx/__fixtures__/nested-multi-row/native-measurements.json',
			import.meta.url,
		),
		'utf8',
	),
) as {
	cases: { name: string; measurements: { lines: { text: string }[] }[] }[];
};
const threeEvidence = JSON.parse(
	readFileSync(
		new NodeURL(
			'../../../core/docx/__fixtures__/nested-three-row/native-measurements.json',
			import.meta.url,
		),
		'utf8',
	),
) as typeof evidence;
const structuralEvidence = JSON.parse(
	readFileSync(
		new NodeURL(
			'../../../core/docx/__fixtures__/nested-structural-rows/native-measurements.json',
			import.meta.url,
		),
		'utf8',
	),
) as typeof evidence;
for (const { native, directory } of [
	...evidence.cases.map((native) => ({ native, directory: 'nested-multi-row' })),
	...threeEvidence.cases.map((native) => ({ native, directory: 'nested-three-row' })),
	...structuralEvidence.cases.map((native) => ({ native, directory: 'nested-structural-rows' })),
])
	it('paints source outer edges once across automatic row fragments: ' + native.name, async () => {
		const loaded = await loadDocx(
			new Uint8Array(
				readFileSync(
					new NodeURL(
						'../../../core/docx/__fixtures__/' + directory + '/' + native.name + '.docx',
						import.meta.url,
					),
				),
			),
		);
		const { element } = renderPrintLayout(
			layoutDocumentModel(loaded.model, {
				widthOf: (text) => text.length * 5,
				lineHeightOf: () => 16,
			}),
		);
		const pages = [...element.querySelectorAll('.dve-print-page')];
		expect(pages).toHaveLength(native.measurements.length);
		pages.forEach((page, index) => {
			const frames = page.querySelectorAll<HTMLElement>('.dve-print-nested-table');
			expect(frames).toHaveLength(1);
			const frame = frames[0]!;
			expect(frame.textContent).toBe('');
			expect(frame.style.pointerEvents).toBe('none');
			const rows = frame.querySelectorAll<HTMLElement>(':scope > .dve-print-row');
			expect(
				parseFloat(rows[0]!.querySelector<HTMLElement>('.dve-print-cell')!.style.borderTopWidth),
			).toBeCloseTo(8 / 3);
			expect(
				parseFloat(
					rows[rows.length - 1]!.querySelector<HTMLElement>('.dve-print-cell')!.style
						.borderBottomWidth,
				),
			).toBeCloseTo(2 / 3);
			const tokens = (page.textContent ?? '').match(/T[012]R\dC\dL\d{2}/g) ?? [];
			expect(tokens.sort()).toEqual(
				native.measurements[index]!.lines.map((line) => line.text.trim())
					.filter((text) => /^T[012]/.test(text))
					.sort(),
			);
			const layer = frame.closest<HTMLElement>('.dve-print-cell-frame-layer')!;
			expect(layer.style.pointerEvents).toBe('none');
			expect(layer.style.overflowY).toBe('clip');
		});
	});
