// @vitest-environment jsdom
import { readFile } from 'node:fs/promises';
import { resolve } from 'node:path';
import { describe, expect, it, vi } from 'vitest';
import { loadDocx, createDocument, type ReviewDisplayMode } from 'ooxml-core/docx';
import { createPrintLayoutController } from './print-layout-view';

describe('Original formatting in Print Layout', () => {
	for (const [area, name] of [
		['review-formatting', 'bold'],
		['review-formatting', 'multiple'],
		['review-paragraph-formatting', 'multiple'],
	])
		it(`matches native ${area} ${name} before appearance without changing the current model`, async () => {
			const load = async (suffix: string) =>
				loadDocx(
					new Uint8Array(
						await readFile(resolve('../core/docx/__fixtures__', area!, `${name}-${suffix}.docx`)),
					),
				);
			const before = (await load('before')).model;
			const current = (await load('tracked')).model;
			const source = structuredClone(current);
			vi.useFakeTimers();
			let mode: ReviewDisplayMode = 'all';
			const controller = createPrintLayoutController(
				document.createElement('div'),
				() => {},
				undefined,
				undefined,
				() => mode,
			);
			try {
				controller.scheduleRelayout(before);
				vi.runAllTimers();
				const expected = controller.element.innerHTML;
				controller.scheduleRelayout(current);
				vi.runAllTimers();
				const final = controller.element.innerHTML;
				expect(final).not.toBe(expected);
				mode = 'original';
				controller.scheduleRelayout(current);
				vi.runAllTimers();
				expect(controller.element.innerHTML).toBe(expected);
				expect(current).toEqual(source);
				mode = 'final';
				controller.scheduleRelayout(current);
				vi.runAllTimers();
				expect(controller.element.innerHTML).toBe(final);
			} finally {
				controller.destroy();
				vi.useRealTimers();
			}
		});

	it('reports missing history and keeps current content rendered', () => {
		const model = createDocument();
		model.blocks = [
			{
				type: 'paragraph',
				id: 'p',
				align: 'center',
				formatRevision: { id: '7', kind: 'paragraphChange', author: 'Ada' },
				runs: [
					{
						text: 'Text',
						bold: true,
						formatRevision: { id: '8', kind: 'formatChange', author: 'Ada' },
					},
					{ text: 'Added', revision: { id: '9', kind: 'insert', author: 'Ada' } },
				],
			},
		];
		vi.useFakeTimers();
		const print = vi.spyOn(window, 'print').mockImplementation(() => {});
		const controller = createPrintLayoutController(
			document.createElement('div'),
			() => {},
			undefined,
			undefined,
			() => 'original',
		);
		try {
			controller.scheduleRelayout(model);
			vi.runAllTimers();
			expect(controller.element.textContent).toContain('Text');
			expect(
				controller
					.approximations()
					.filter((message) => message.startsWith('Original formatting unavailable')),
			).toHaveLength(2);
			expect(controller.approximations()).toContain(
				'Print Layout retains tracked text and paragraph marks; this review display projects formatting only.',
			);
			const note = vi.fn();
			controller.print(model, note);
			expect(note).toHaveBeenCalledTimes(3);
			expect(print).toHaveBeenCalledOnce();
		} finally {
			controller.destroy();
			print.mockRestore();
			vi.useRealTimers();
		}
	});
});
