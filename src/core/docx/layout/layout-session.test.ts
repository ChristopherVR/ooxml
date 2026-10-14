import { expect, it, vi } from 'vitest';
import { createLayoutSession } from './layout-session';
import { layoutDocument } from './layout';
import { createFakeMeasurer } from './measure';
import type { LayoutDocumentInput } from './input';

function input(text = 'First paragraph'): LayoutDocumentInput {
	return {
		sections: [
			{
				page: {
					widthPx: 300,
					heightPx: 200,
					marginTopPx: 10,
					marginRightPx: 10,
					marginBottomPx: 10,
					marginLeftPx: 10,
				},
				blocks: [text, 'Second unchanged paragraph'].map((text, i) => ({
					kind: 'paragraph',
					id: String(i),
					runs: [{ text }],
				})),
			},
		],
	};
}

it('remeasures changed paragraphs while matching uncached pagination', () => {
	const measurer = createFakeMeasurer();
	const width = vi.spyOn(measurer, 'widthOf');
	try {
		const session = createLayoutSession(measurer);
		session.layoutDocument(input());
		width.mockClear();
		const unchanged = session.layoutDocument(input());
		expect(width).not.toHaveBeenCalled();
		const changed = session.layoutDocument(input('Edited text'));
		expect(width).toHaveBeenCalled();
		expect(width.mock.calls.some(([text]) => text === 'unchanged')).toBe(false);
		expect(changed).toEqual(layoutDocument(input('Edited text'), createFakeMeasurer()));
		unchanged.pages[0]!.columns[0]!.blocks.length = 0;
		expect(session.layoutDocument(input())).toEqual(layoutDocument(input(), createFakeMeasurer()));
	} finally {
		width.mockRestore();
	}
});

it('replays warnings and invalidates width and formatting changes', () => {
	const session = createLayoutSession(createFakeMeasurer());
	const document = input('x'.repeat(200));
	const first = session.layoutDocument(document);
	const second = session.layoutDocument(document);
	second.pages[0]!.columns[0]!.blocks[0]!.blockId = 'corrupted';
	const restored = session.layoutDocument(document);
	const fresh = layoutDocument(document, createFakeMeasurer());
	// Measured paragraph copies cannot leak mutations from a returned document into later passes.
	const block = first.pages[0]!.columns[0]!.blocks[0]!;
	if (block.kind === 'paragraph') block.lines[0]!.fragments[0]!.text = 'changed';
	expect(restored).toEqual(fresh);
	expect(second.approximations).toEqual(fresh.approximations);
	document.sections[0]!.page.widthPx = 500;
	expect(session.layoutDocument(document)).toEqual(layoutDocument(document, createFakeMeasurer()));
});
