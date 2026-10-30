import { describe, expect, it } from 'vitest';
import {
	pdfOptionsFor,
	pdfOptionsPanel,
	pdfReportLines,
	readPdfChoice,
	showPdfReport,
} from './viewer-pdf-options';

const panel = () => {
	const host = document.createElement('div');
	host.append(pdfOptionsPanel(document));
	return host;
};
const choose = (host: HTMLElement, scope: string, from = '1', to = '3') => {
	const select = host.querySelector<HTMLSelectElement>('[data-pdf-scope]')!;
	select.value = scope;
	select.dispatchEvent(new Event('change'));
	host.querySelector<HTMLInputElement>('[data-pdf-from]')!.value = from;
	host.querySelector<HTMLInputElement>('[data-pdf-to]')!.value = to;
};

describe('PDF options panel', () => {
	it('defaults to all pages and shows the range inputs only for a range', () => {
		const host = panel();
		expect(readPdfChoice(host, 4)).toEqual({ scope: 'all' });
		const range = host.querySelector<HTMLElement>('[data-pdf-range]')!;
		expect(range.hidden).toBe(true);
		choose(host, 'range');
		expect(range.hidden).toBe(false);
	});

	it('reads the current page and clamps a range to the foreground pages', () => {
		const host = panel();
		choose(host, 'current');
		expect(readPdfChoice(host, 4)).toEqual({ scope: 'current' });
		choose(host, 'range', '0', '99');
		expect(readPdfChoice(host, 4)).toEqual({ scope: 'range', from: 1, to: 4 });
		choose(host, 'range', '3', '2');
		expect(readPdfChoice(host, 4)).toEqual({ scope: 'range', from: 3, to: 3 });
	});

	it('maps a choice to export options', () => {
		expect(pdfOptionsFor({ scope: 'all' }, 2)).toEqual({});
		expect(pdfOptionsFor({ scope: 'current' }, 2)).toEqual({ scope: 'current', pageIndex: 2 });
		expect(pdfOptionsFor({ scope: 'range', from: 2, to: 3 }, 0)).toEqual({ from: 2, to: 3 });
	});

	it('lists what each page lost and which fonts were embedded', () => {
		const summary = {
			pages: [
				{ index: 0, name: 'Plan', mode: 'vector' as const, fallback: [], degraded: ['Glow x.'] },
				{
					index: 1,
					name: 'Chart',
					mode: 'picture' as const,
					fallback: ['Unsupported content: embedded chart.'],
					degraded: [],
				},
			],
			embeddedFonts: ['Arial'],
			refusedFonts: ['Locked'],
		};
		const lines = pdfReportLines(summary);
		expect(lines[0]).toBe('Plan: drawn as vectors with selectable text.');
		expect(lines).toContain('Plan: Glow x.');
		expect(lines.join('\n')).toMatch(
			/Chart: saved as a picture \(Unsupported content: embedded chart/,
		);
		expect(lines.join('\n')).toMatch(/Fonts embedded: Arial/);
		expect(lines.join('\n')).toMatch(/Fonts not embedded.*Locked/);
		const host = panel();
		showPdfReport(host, summary);
		const report = host.querySelector<HTMLElement>('[data-pdf-report]')!;
		expect(report.hidden).toBe(false);
		expect(report.querySelectorAll('li').length).toBe(lines.length);
		showPdfReport(host, undefined);
		expect(report.hidden).toBe(true);
	});
});
