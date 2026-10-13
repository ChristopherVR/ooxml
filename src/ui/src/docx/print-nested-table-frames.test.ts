// @vitest-environment jsdom
import { readFileSync } from 'node:fs';
import { expect, it } from 'vitest';
import { loadDocx } from 'ooxml-core/docx';
import { layoutDocumentModel } from 'ooxml-core/docx/layout';
import { renderPrintLayout } from './print-layout';
it.each(['short', 'two-rows', 'centered', 'split', 'align-top', 'align-center', 'align-bottom'])(
	'paints nested source frames without duplicating ink: %s',
	async (name) => {
		const loaded = await loadDocx(
			new Uint8Array(
				readFileSync(
					new URL(
						'../../../core/docx/__fixtures__/nested-table-frames/nested-frame-' + name + '.docx',
						import.meta.url,
					),
				),
			),
		);
		const layout = layoutDocumentModel(loaded.model, {
			widthOf: (text) => text.length * 5,
			lineHeightOf: () => 16,
		});
		const { element } = renderPrintLayout(layout);
		const frames = [...element.querySelectorAll<HTMLElement>('.dve-print-nested-table')];
		expect(frames.length).toBeGreaterThan(0);
		for (const frame of frames) {
			expect(frame.getAttribute('aria-hidden')).toBe('true');
			expect(frame.style.pointerEvents).toBe('none');
			expect(frame.textContent).toBe('');
			expect(frame.querySelector<HTMLElement>('.dve-print-cell')?.style.borderLeft).toContain(
				'rgb(255, 0, 0)',
			);
			expect(frame.querySelector<HTMLElement>('.dve-print-cell')?.style.background).toBe(
				'rgb(255, 255, 0)',
			);
		}
		const text = element.textContent ?? '';
		expect(text.match(/T1R0C0L00/g)).toHaveLength(1);
		expect(text.match(/T1R0C1L00/g)).toHaveLength(1);
		if (name === 'split') expect(frames).toHaveLength(2);
	},
);
