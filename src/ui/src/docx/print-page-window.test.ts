// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';
import { PrintPageWindow } from './print-page-window';

function windowOf(count: number) {
	const root = document.createElement('div');
	const pages = new PrintPageWindow(root, true);
	const sheets = Array.from({ length: count }, () => document.createElement('div'));
	for (const sheet of sheets)
		pages.add(sheet, () => {
			sheet.append(document.createElement('p'));
			return () => undefined;
		});
	return { pages, sheets };
}
const mounted = (sheets: HTMLElement[]) =>
	sheets.flatMap((sheet, index) => (sheet.childElementCount ? [index] : []));

describe('PrintPageWindow', () => {
	it('mounts the window around the centre and drops pages that leave it', () => {
		const { pages, sheets } = windowOf(14);
		pages.show(0);
		expect(mounted(sheets)).toEqual([0, 1, 2]);
		pages.show(9);
		expect(mounted(sheets)).toEqual([7, 8, 9, 10, 11]);
	});

	it('keeps pinned pages mounted while the window moves away', () => {
		const { pages, sheets } = windowOf(14);
		pages.show(0, [0, undefined]);
		pages.show(9, [0]);
		expect(mounted(sheets)).toEqual([0, 7, 8, 9, 10, 11]);
		pages.show(9, 3);
		expect(mounted(sheets)).toEqual([3, 7, 8, 9, 10, 11]);
	});
});
