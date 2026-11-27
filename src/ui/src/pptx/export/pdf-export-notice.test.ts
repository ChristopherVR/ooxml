import { afterEach, describe, expect, it, vi } from 'vitest';

import { PDF_SVG_NOTE_CODES } from 'ooxml-core/pdf/write';
import { translationsEn } from '../i18n/translations-en';
import { createTranslator } from '../i18n/translator';
import * as pdfVectorExportEntry from './pdf-vector-export-entry';
import {
	distinctPdfDiagnostics,
	PDF_EXPORT_DIAGNOSTIC_CODES,
	pdfDiagnosticCategory,
	pdfDiagnosticMessageKey,
	type PdfExportDiagnostic,
} from './pdf-export-diagnostics';
import {
	PDF_EXPORT_NOTICE_ID,
	pdfDiagnosticMessage,
	pdfExportNoticeToast,
	summarizePdfExport,
	type PdfExportReportSource,
} from './pdf-export-notice';
import { runPdfExport } from './pdf-export-run';

const font = (family: string, substitute = 'Helvetica'): PdfExportDiagnostic => ({
	code: 'fontStandard',
	params: { family, substitute },
});

function report(slides: PdfExportReportSource['slides']): PdfExportReportSource {
	return { slides, diagnostics: distinctPdfDiagnostics(slides.flatMap((s) => s.diagnostics)) };
}

describe('PDF export diagnostic codes', () => {
	it('gives every code a category and an English message, with no sentence in the code', () => {
		for (const code of PDF_EXPORT_DIAGNOSTIC_CODES) {
			expect(code).toMatch(/^[a-z][A-Za-z0-9]+$/u);
			expect(pdfDiagnosticCategory(code)).toBeTruthy();
			expect(translationsEn[pdfDiagnosticMessageKey(code)], code).toBeTruthy();
		}
		for (const code of PDF_SVG_NOTE_CODES) expect(pdfDiagnosticCategory(code)).toBe('simplified');
		expect(translationsEn['pptx.pdfExport.diagnostic.fontStandardDefaultFamily']).toBeTruthy();
		expect(translationsEn['pptx.pdfExport.diagnostic.fontSubstitutedDefaultFamily']).toBeTruthy();
	});

	it('dedupes on code and params, keeping the first seen first', () => {
		expect(
			distinctPdfDiagnostics([
				{ code: 'shadowAsPicture' },
				font('Calibri'),
				{ code: 'shadowAsPicture' },
				font('Calibri'),
				font('Aptos'),
			]),
		).toEqual([{ code: 'shadowAsPicture' }, font('Calibri'), font('Aptos')]);
	});

	it('maps a font diagnostic without a family to the default-font message', () => {
		expect(pdfDiagnosticMessage(font(''))).toEqual({
			key: 'pptx.pdfExport.diagnostic.fontStandardDefaultFamily',
			params: { substitute: 'Helvetica' },
		});
		expect(pdfDiagnosticMessage({ code: 'blendAsPicture' })).toEqual({
			key: 'pptx.pdfExport.diagnostic.blendAsPicture',
		});
	});
});

describe('pdfExportNoticeToast', () => {
	it('counts pictures, whole-slide pictures, distinct fonts, omissions and simplifications', () => {
		const result = report([
			{
				mode: 'vector',
				pictures: 2,
				diagnostics: [{ code: 'shadowAsPicture' }, font('Calibri'), { code: 'glowApproximated' }],
			},
			{ mode: 'vector', pictures: 1, diagnostics: [font('Calibri'), { code: 'pictureMissing' }] },
			{ mode: 'picture', diagnostics: [{ code: 'slideTooMuchText' }, { code: 'searchableText' }] },
		]);
		expect(summarizePdfExport(result)).toEqual({
			pictures: 3,
			slidesAsPictures: 1,
			fonts: 1,
			omitted: 1,
			simplified: 1,
		});
	});

	it('summarises in the viewer language with a details list', () => {
		const toast = pdfExportNoticeToast(
			report([
				{
					mode: 'vector',
					pictures: 3,
					diagnostics: [{ code: 'filterAsPicture' }, font('Calibri')],
				},
			]),
		)!;
		expect(toast).toMatchObject({ id: 'pdf-export-report', severity: 'info' });
		const t = createTranslator();
		const parts = toast.parts!.map((part) => t(part.key, part.params));
		expect(t(toast.messageKey, { parts: parts.join(t('pptx.compatibility.partSeparator')) })).toBe(
			'PDF saved with approximations: 3 elements were saved as pictures; 1 font was substituted.',
		);
		expect(toast.details!.map((line) => t(line.key, line.params))).toEqual([
			'Soft edges, blur, glow and other filter effects are drawn as pictures.',
			expect.stringContaining('Text in "Calibri" uses a standard PDF font (Helvetica)'),
		]);
	});

	it('shows no notice when nothing was approximated', () => {
		expect(pdfExportNoticeToast(report([{ mode: 'vector', pictures: 0, diagnostics: [] }]))).toBe(
			undefined,
		);
		expect(pdfExportNoticeToast(undefined)).toBeUndefined();
	});
});

describe('runPdfExport notice', () => {
	afterEach(() => vi.restoreAllMocks());

	it('never notifies for the picture export and dismisses any earlier notice', async () => {
		const notify = vi.fn();
		const dismissNotice = vi.fn();
		const pictures = vi.fn(() => Promise.resolve());
		await runPdfExport({
			settings: { mode: 'pictures', includeHiddenSlides: false },
			vector: {} as never,
			pictures,
			save: vi.fn(),
			notify,
			dismissNotice,
		});
		expect(pictures).toHaveBeenCalledOnce();
		expect(notify).not.toHaveBeenCalled();
		expect(dismissNotice).toHaveBeenCalledWith(PDF_EXPORT_NOTICE_ID);
	});

	it('dismisses an earlier notice when a vector export completes cleanly with no approximations', async () => {
		vi.spyOn(pdfVectorExportEntry, 'exportSlidesVectorPdf').mockResolvedValue({
			blob: new Blob(),
			bytes: new Uint8Array(),
			slides: [{ index: 0, mode: 'vector', pictures: 0, diagnostics: [] }],
			diagnostics: [],
		} as never);
		const notify = vi.fn();
		const dismissNotice = vi.fn();
		const save = vi.fn();
		await runPdfExport({
			settings: { mode: 'vector', includeHiddenSlides: false },
			vector: {} as never,
			pictures: vi.fn(() => Promise.resolve()),
			save,
			notify,
			dismissNotice,
		});
		expect(save).toHaveBeenCalledOnce();
		expect(notify).not.toHaveBeenCalled();
		expect(dismissNotice).toHaveBeenCalledWith(PDF_EXPORT_NOTICE_ID);
	});
});
