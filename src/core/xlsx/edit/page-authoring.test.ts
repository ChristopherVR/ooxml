import { expect, it } from 'vitest';
import { createEditSession, createWorkbook, loadXlsx, saveXlsx } from '../index';

it('round-trips fit-to-page, percentage scaling, print titles and header/footer text', async () => {
	const workbook = createWorkbook();
	const session = createEditSession(workbook);
	session.setPageSetup(0, { fitToWidth: 2, fitToHeight: 0 });
	session.setPrintTitles(0, { rows: { start: 0, end: 1 }, columns: { start: 0, end: 0 } });
	session.setSheetView(0, { mode: 'pageLayout' });
	session.setPageSetup(0, { header: '&CForecast &P', footer: '&R&A' });
	const fitted = await loadXlsx(await saveXlsx(workbook));
	expect(fitted.sheets[0]?.pageSetup).toMatchObject({
		fitToWidth: 2,
		fitToHeight: 0,
		header: '&CForecast &P',
		footer: '&R&A',
	});
	expect(fitted.sheets[0]?.view.mode).toBe('pageLayout');
	expect(fitted.definedNames).toContainEqual({
		name: '_xlnm.Print_Titles',
		formula: 'Sheet1!$1:$2,Sheet1!$A:$A',
		localSheet: 0,
	});
	expect(session.undo()).toBe(true);
	expect(workbook.sheets[0]?.pageSetup?.header).toBeUndefined();
	expect(session.undo()).toBe(true);
	expect(workbook.sheets[0]?.view.mode).toBeUndefined();
	expect(session.undo()).toBe(true);
	expect(workbook.definedNames).toHaveLength(0);

	const fitSession = createEditSession(fitted);
	fitSession.setPageSetup(0, { scale: 75, fitToWidth: undefined, fitToHeight: undefined });
	const scaled = await loadXlsx(await saveXlsx(fitted));
	expect(scaled.sheets[0]?.pageSetup?.scale).toBe(75);
	expect(scaled.sheets[0]?.pageSetup?.fitToWidth).toBeUndefined();
	expect(fitSession.undo()).toBe(true);
	expect(fitted.sheets[0]?.pageSetup?.fitToWidth).toBe(2);
});

it('round-trips a worksheet background and removes its package part on deletion', async () => {
	const workbook = createWorkbook();
	const session = createEditSession(workbook);
	const bytes = Uint8Array.from([137, 80, 78, 71]);
	session.setSheetBackground(0, { bytes, contentType: 'image/png' });
	expect(session.undo()).toBe(true);
	expect(workbook.sheets[0]?.backgroundPicture).toBeUndefined();
	expect(session.redo()).toBe(true);
	const reopened = await loadXlsx(await saveXlsx(workbook));
	const picture = reopened.sheets[0]?.backgroundPicture;
	expect(picture?.contentType).toBe('image/png');
	expect(reopened.source?.parts.get(picture?.partName ?? '')).toEqual(bytes);
	expect(reopened.sheets[0]?.drawings).toEqual([]);

	const edit = createEditSession(reopened);
	edit.setSheetBackground(0, undefined);
	const deleted = await loadXlsx(await saveXlsx(reopened));
	expect(deleted.sheets[0]?.backgroundPicture).toBeUndefined();
	expect(deleted.source?.parts.has(picture?.partName ?? '')).toBe(false);
	expect(edit.undo()).toBe(true);
	expect(reopened.sheets[0]?.backgroundPicture).toEqual(picture);
});

it('round-trips a legacy note beside a threaded comment', async () => {
	const workbook = createWorkbook();
	const session = createEditSession(workbook);
	session.setComment(0, { row: 0, col: 0 }, 'Legacy note', 'Author');
	session.setComment(0, { row: 0, col: 1 }, 'Thread root', 'Author', true);
	const reopened = await loadXlsx(await saveXlsx(workbook));
	expect(reopened.sheets[0]?.comments).toHaveLength(2);
	expect(reopened.sheets[0]?.comments[0]).toMatchObject({ text: 'Legacy note' });
	expect(reopened.sheets[0]?.comments[0]?.replies).toBeUndefined();
	expect(reopened.sheets[0]?.comments[1]).toMatchObject({ text: 'Thread root', replies: [] });
	expect(session.undo()).toBe(true);
	expect(session.undo()).toBe(true);
	expect(workbook.sheets[0]?.comments).toHaveLength(0);
});
