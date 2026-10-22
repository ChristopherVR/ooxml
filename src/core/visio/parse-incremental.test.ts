import { describe, expect, it } from 'vitest';
import { createVsdx, editVsdx, mergeVisioPages, parseVsdx, parseVsdxPages } from './index';
import type { VisioDiagnostic, VisioDocument } from './model';

const rect = (pageId: string, shapeId: string, x: number) =>
	({ type: 'create-rectangle', pageId, shapeId, x, y: 2, width: 1, height: 1 }) as const;

async function twoPages() {
	let bytes = (await editVsdx(await createVsdx(), [rect('0', '1', 1), rect('0', '2', 3)])).bytes;
	bytes = (
		await editVsdx(bytes, [{ type: 'insert-page', pageId: '5', afterPageId: '0', name: 'Two' }])
	).bytes;
	return (await editVsdx(bytes, [rect('5', '1', 2)])).bytes;
}

describe('reading some pages', () => {
	it('lists only the requested pages and skips the list-level extras', async () => {
		const bytes = await twoPages();
		const partial = await parseVsdx(bytes, { pages: ['5'] });
		expect(partial.pages.map((p) => p.id)).toEqual(['5']);
		const whole = await parseVsdx(bytes);
		expect(partial.pages[0]).toEqual(whole.pages.find((p) => p.id === '5'));
	});

	it('merges to the document a whole read of the edited package gives', async () => {
		const before = await twoPages();
		const previous = await parseVsdx(before);
		const after = (
			await editVsdx(before, [{ type: 'move-shape', pageId: '5', shapeId: '1', x: 6, y: 2 }])
		).bytes;
		const merged = await parseVsdxPages(after, previous, ['5']);
		expect(merged).toEqual(await parseVsdx(after));
		// The page that did not change is the very object that was read before.
		expect(merged!.pages.find((p) => p.id === '0')).toBe(previous.pages.find((p) => p.id === '0'));
		expect(merged!.pages.find((p) => p.id === '5')).not.toBe(
			previous.pages.find((p) => p.id === '5'),
		);
	});

	it('merges edits of two pages and added and removed shapes', async () => {
		const before = await twoPages();
		const previous = await parseVsdx(before);
		const after = (
			await editVsdx(before, [
				rect('0', '9', 5),
				{ type: 'delete-shape', pageId: '5', shapeId: '1' },
				{ type: 'replace-plain-text', pageId: '0', shapeId: '1', text: 'hi' },
			])
		).bytes;
		const merged = await parseVsdxPages(after, previous, ['0', '5']);
		expect(merged).toEqual(await parseVsdx(after));
	});

	it('declines when the pages do not line up with the earlier document', async () => {
		const bytes = await twoPages();
		const previous = await parseVsdx(bytes);
		expect(await parseVsdxPages(bytes, previous, ['nope'])).toBeUndefined();
		const partial = await parseVsdx(bytes, { pages: ['5'] });
		expect(mergeVisioPages(previous, partial, ['5', '0'])).toBeUndefined();
		expect(mergeVisioPages(previous, partial, ['5'], { maxShapes: 1 })).toBeUndefined();
	});
});

describe('splicing diagnostics', () => {
	const note = (code: string, pageId?: string): VisioDiagnostic => ({
		code,
		severity: 'warning',
		message: code,
		...(pageId === undefined ? {} : { pageId }),
	});
	async function docs(previous: VisioDiagnostic[], partial: VisioDiagnostic[]) {
		const bytes = await twoPages();
		const base = await parseVsdx(bytes);
		const prev: VisioDocument = { ...base, diagnostics: previous };
		const part = await parseVsdx(bytes, { pages: ['5'] });
		return { prev, part: { ...part, diagnostics: partial } };
	}

	it('replaces the block of a page in place and keeps the rest in order', async () => {
		const { prev, part } = await docs(
			[
				note('a'),
				note('p0', '0'),
				note('p5a', '5'),
				note('p5b', '5'),
				note('background-cycle', '5'),
				note('z'),
			],
			[note('doc'), note('n1', '5'), note('n2', '5'), note('n3', '5')],
		);
		const merged = mergeVisioPages(prev, part, ['5'])!;
		expect(merged.diagnostics.map((d) => d.code)).toEqual([
			'a',
			'p0',
			'n1',
			'n2',
			'n3',
			'background-cycle',
			'z',
		]);
	});

	it('keeps the background checks of the page list', async () => {
		const { prev, part } = await docs(
			[note('p5', '5'), note('background-cycle', '5'), note('z')],
			[],
		);
		const merged = mergeVisioPages(prev, part, ['5'])!;
		expect(merged.diagnostics.map((d) => d.code)).toEqual(['background-cycle', 'z']);
	});

	it('declines when a new diagnostic cannot be placed or limits may apply', async () => {
		const added = await docs([note('a'), note('z')], [note('n', '5')]);
		expect(mergeVisioPages(added.prev, added.part, ['5'])).toBeUndefined();
		const cut = await docs([note('p', '5'), note('diagnostics-truncated')], []);
		expect(mergeVisioPages(cut.prev, cut.part, ['5'])).toBeUndefined();
		const split = await docs([note('p', '5'), note('a'), note('q', '5')], []);
		expect(mergeVisioPages(split.prev, split.part, ['5'])).toBeUndefined();
	});
});
