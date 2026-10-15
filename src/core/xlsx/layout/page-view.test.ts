import { expect, it } from 'vitest';
import { createWorkbook } from '../workbook';
import { createEditSession } from '../edit/session';
import { parseRange } from '../address';
import { loadXlsx } from '../read/index';
import { saveXlsx } from '../write/index';
import { createGridMetrics } from './metrics';
import { pageOfCell, pageViewCellRect, pageViewHit, pageViewLayout } from './page-view';
import { pageBreakPreview, planBreakDrag, resizePrintArea } from './page-break-view';
import {
	composeHeaderFooter,
	headerFooterKey,
	patchHeaderFooterSection,
	splitHeaderFooter,
} from './header-footer-edit';

const setup = () => {
	const wb = createWorkbook();
	const edit = createEditSession(wb);
	for (let r = 0; r < 10; r++) for (let c = 0; c < 4; c++) edit.setCellInput(0, r, c, `${r}:${c}`);
	return { wb, edit };
};

it('places pages in a grid with margins, header and footer bands', () => {
	const { wb, edit } = setup();
	edit.setPageBreak(0, 'row', 5);
	edit.setPageBreak(0, 'col', 2);
	const layout = pageViewLayout(wb, 0);
	expect(layout.pages.map((p) => [p.slotRow, p.slotCol])).toEqual([
		[0, 0],
		[1, 0],
		[0, 1],
		[1, 1],
	]);
	const first = layout.pages[0]!;
	expect(first.paper).toMatchObject({ x: 0, y: 0, width: 816, height: 1056 });
	expect(first.content.x).toBeCloseTo(0.7 * 96);
	expect(first.header.y).toBeCloseTo(0.3 * 96);
	expect(first.footer.y).toBeCloseTo(1056 - 0.75 * 96);
	expect(layout.pages[2]!.paper.x).toBe(816 + 24);
});

it('hits cells, header sections and margins, and gives cell rectangles', () => {
	const { wb } = setup();
	const layout = pageViewLayout(wb, 0);
	const metrics = createGridMetrics(wb.sheets[0]!, { zoom: 100 });
	const page = layout.pages[0]!;
	const rect = pageViewCellRect(layout, metrics, 1, 2, 1)!;
	expect(pageViewHit(layout, metrics, rect.x + 2, rect.y + 2)).toEqual({
		kind: 'cell',
		page: 1,
		row: 2,
		col: 1,
	});
	expect(pageViewHit(layout, metrics, page.header.x + 5, page.header.y + 2)).toMatchObject({
		kind: 'header',
		section: 'left',
	});
	expect(
		pageViewHit(layout, metrics, page.footer.x + page.footer.width - 5, page.footer.y + 2),
	).toMatchObject({ kind: 'footer', section: 'right' });
	expect(pageViewHit(layout, metrics, 2, 2)).toEqual({ kind: 'margin', page: 1 });
	expect(pageViewHit(layout, metrics, -5, 5)).toBeUndefined();
	expect(pageOfCell(layout, 3, 3)).toBe(1);
});

it('shows one blank page for an empty sheet and honours landscape', () => {
	const wb = createWorkbook();
	createEditSession(wb).setPageSetup(0, { orientation: 'landscape', paperSize: 9 });
	const layout = pageViewLayout(wb, 0);
	expect(layout.pages).toHaveLength(1);
	expect(layout.pages[0]!.blank).toBe(true);
	expect(layout.pages[0]!.paper.width).toBeGreaterThan(layout.pages[0]!.paper.height);
});

it('lists automatic and manual break lines and plans drags', () => {
	const { wb, edit } = setup();
	edit.setPageBreak(0, 'row', 5);
	const preview = pageBreakPreview(wb, 0);
	expect(preview.area).toEqual(parseRange('A1:D10'));
	expect(preview.lines).toEqual([{ axis: 'row', before: 5, manual: true }]);
	const line = preview.lines[0]!;
	expect(planBreakDrag(preview, line, 7)).toEqual([
		{ axis: 'row', before: 5, present: false },
		{ axis: 'row', before: 7, present: true },
	]);
	expect(planBreakDrag(preview, line, 5)).toEqual([]);
	expect(planBreakDrag(preview, line, 0)).toEqual([]);
	expect(planBreakDrag(preview, line, 11)).toEqual([]);
	const auto = { axis: 'col', before: 2, manual: false } as const;
	expect(planBreakDrag(preview, auto, 3)).toEqual([{ axis: 'col', before: 3, present: true }]);
	for (const e of planBreakDrag(preview, line, 7))
		edit.setPageBreak(0, e.axis, e.before, e.present);
	expect(pageBreakPreview(wb, 0).lines).toEqual([{ axis: 'row', before: 7, manual: true }]);
	edit.undo();
	edit.undo();
	expect(pageBreakPreview(wb, 0).lines.map((l) => l.before)).toEqual([5]);
});

it('marks automatic breaks and ignores drags under fit to pages', () => {
	const { wb, edit } = setup();
	edit.setPageSetup(0, { scale: 400 });
	const auto = pageBreakPreview(wb, 0);
	expect(auto.lines.length).toBeGreaterThan(0);
	expect(auto.lines.every((l) => !l.manual)).toBe(true);
	edit.setPageSetup(0, { scale: undefined, fitToWidth: 1, fitToHeight: 1 });
	const fit = pageBreakPreview(wb, 0);
	expect(fit.fit).toBe(true);
	expect(planBreakDrag(fit, { axis: 'row', before: 2, manual: false }, 3)).toEqual([]);
});

it('resizes the print area by edge', () => {
	const area = parseRange('B2:D6')!;
	expect(resizePrintArea(area, 'bottom', 9)).toEqual(parseRange('B2:D10')!);
	expect(resizePrintArea(area, 'left', 0)).toEqual(parseRange('A2:D6')!);
	expect(resizePrintArea(area, 'top', 20)).toEqual(parseRange('B6:D6')!);
	expect(resizePrintArea(area, 'right', -1)).toBeUndefined();
});

it('splits, edits and composes header sections', () => {
	expect(splitHeaderFooter('&LLeft&CPage &P of &N&R&D')).toEqual({
		left: 'Left',
		center: 'Page &P of &N',
		right: '&D',
	});
	expect(splitHeaderFooter('Plain &&')).toEqual({ left: '', center: 'Plain &&', right: '' });
	expect(splitHeaderFooter('&L&"Arial,Bold"&12Hi&RR').left).toBe('&"Arial,Bold"&12Hi');
	expect(composeHeaderFooter({ left: 'a', right: 'b' })).toBe('&La&Rb');
	expect(headerFooterKey({ differentFirst: true }, 'header', 1, 1)).toBe('firstHeader');
	expect(headerFooterKey({ differentOddEven: true }, 'footer', 2, 2)).toBe('evenFooter');
	expect(patchHeaderFooterSection({ header: '&LA' }, 'header', 'center', '&F')).toEqual({
		header: '&LA&C&F',
	});
	expect(patchHeaderFooterSection({ header: '&LA' }, 'header', 'left', '')).toEqual({
		header: undefined,
	});
});

it('round-trips the sheet view mode, header/footer and breaks', async () => {
	const { wb, edit } = setup();
	edit.setSheetView(0, { mode: 'pageBreakPreview' });
	edit.setPageSetup(0, { header: '&LTitle&C&A', footer: '&CPage &P of &N' });
	edit.setPageBreak(0, 'row', 4);
	let loaded = await loadXlsx(await saveXlsx(wb));
	expect(loaded.sheets[0]!.view.mode).toBe('pageBreakPreview');
	expect(loaded.sheets[0]!.pageSetup).toMatchObject({
		header: '&LTitle&C&A',
		footer: '&CPage &P of &N',
	});
	expect(pageBreakPreview(loaded, 0).lines).toEqual([{ axis: 'row', before: 4, manual: true }]);
	createEditSession(loaded).setSheetView(0, { mode: 'pageLayout' });
	loaded = await loadXlsx(await saveXlsx(loaded));
	expect(loaded.sheets[0]!.view.mode).toBe('pageLayout');
	createEditSession(loaded).setSheetView(0, { mode: undefined });
	loaded = await loadXlsx(await saveXlsx(loaded));
	expect(loaded.sheets[0]!.view.mode).toBeUndefined();
});
