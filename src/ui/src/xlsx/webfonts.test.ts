// @vitest-environment jsdom
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createWorkbook, putCell, type Workbook } from 'ooxml-core/xlsx';
import { resetWebfontSessionCache } from '../fonts/webfonts';
import { XlsxWebfonts } from './webfonts';

const hosts: HTMLElement[] = [];

function setup(workbook: Workbook | undefined = createWorkbook()) {
	const host = document.createElement('div');
	document.body.append(host);
	hosts.push(host);
	const remeasure = vi.fn();
	const fonts = new XlsxWebfonts({ host, workbook: () => workbook, remeasure });
	return { fonts, remeasure, workbook, link: () => document.getElementById(fonts.linkId) };
}

const settle = () => new Promise((resolve) => setTimeout(resolve, 30));

afterEach(() => {
	for (const host of hosts.splice(0)) host.remove();
	for (const link of document.head.querySelectorAll('link')) link.remove();
	resetWebfontSessionCache();
	vi.restoreAllMocks();
});

describe('XlsxWebfonts', () => {
	it('is off by default: no link, no wait', async () => {
		const { fonts, workbook, link } = setup();
		expect(fonts.option).toBeUndefined();
		expect(fonts.base()).toBeNull();
		const readiness = await fonts.prepare(workbook!);
		expect(readiness).toEqual({ href: null, families: [], unavailable: [], timedOut: false });
		expect(link()).toBeNull();
	});

	it('makes no link for false or an invalid baseUrl', async () => {
		const warn = vi.spyOn(console, 'warn').mockImplementation(() => undefined);
		const { fonts, workbook, link } = setup();
		fonts.option = false;
		await fonts.prepare(workbook!);
		expect(link()).toBeNull();
		fonts.option = { baseUrl: 'ftp://fonts.example/css2?x=1' };
		await fonts.prepare(workbook!);
		expect(link()).toBeNull();
		expect(warn).toHaveBeenCalledOnce();
	});

	it('links the workbook fonts (as their clones) to the given base', async () => {
		const workbook = createWorkbook();
		const sheet = workbook.sheets[0]!;
		workbook.styles.push({ ...workbook.styles[0]!, font: { name: 'Arial', bold: true } });
		putCell(sheet, 0, 0, { value: 'Total', styleId: 1 });
		const { fonts, link } = setup(workbook);
		fonts.option = { baseUrl: 'https://fonts.example/css2' };
		const readiness = await fonts.prepare(workbook);
		const href = link()?.getAttribute('href') ?? '';
		expect(href.startsWith('https://fonts.example/css2?')).toBe(true);
		expect(decodeURIComponent(href)).toContain('family=Carlito:');
		expect(decodeURIComponent(href)).toContain('family=Arimo:');
		expect(readiness.families).toEqual(expect.arrayContaining(['Carlito', 'Arimo']));
		// Turning it off removes the link.
		fonts.option = false;
		fonts.shown(workbook);
		expect(link()).toBeNull();
	});

	it('measures again once a shown workbook has its faces, and not twice for a waited one', async () => {
		const { fonts, remeasure, workbook } = setup();
		fonts.option = { baseUrl: 'https://fonts.example/css2' };
		fonts.shown(workbook);
		await settle();
		expect(remeasure).toHaveBeenCalledTimes(1);
		// Shown again after the wait settled (a `load()` waits before showing): nothing to redo.
		fonts.shown(workbook);
		await settle();
		expect(remeasure).toHaveBeenCalledTimes(1);
	});

	it('measures again when faces arrive later, coalesced, only while listening', async () => {
		const target = new EventTarget();
		Object.defineProperty(document, 'fonts', { configurable: true, value: target });
		try {
			const { fonts, remeasure } = setup();
			fonts.connect();
			target.dispatchEvent(new Event('loadingdone'));
			target.dispatchEvent(new Event('loadingdone'));
			await settle();
			expect(remeasure).toHaveBeenCalledTimes(1);
			fonts.disconnect();
			target.dispatchEvent(new Event('loadingdone'));
			await settle();
			expect(remeasure).toHaveBeenCalledTimes(1);
		} finally {
			delete (document as { fonts?: unknown }).fonts;
		}
	});
});
