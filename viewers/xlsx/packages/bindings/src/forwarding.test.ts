// @vitest-environment jsdom
// Regression tests: props are forwarded only when they change, controlled parents hear the
// user's changes, and a `src` fetch is dropped once something newer arrives.
import { createWorkbook } from '@christophervr/xlsx-core';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { XlsxEditorElement } from 'xlsx-web-component';
import { nameFromUrl, sameThemeColors } from './forwarding';
import { mountEditor, type EditorOptions } from './index';
import { emit } from './test-support';

afterEach(() => {
	document.body.replaceChildren();
	vi.restoreAllMocks();
	vi.unstubAllGlobals();
});
const newHost = () => {
	const host = document.createElement('div');
	document.body.append(host);
	return host;
};
const flush = () => new Promise((resolve) => setTimeout(resolve, 0));

describe('prop forwarding', () => {
	it('keeps changes the user made inside the editor across unrelated re-renders', () => {
		const props = (): EditorOptions => ({
			readOnly: false,
			locale: 'en',
			theme: 'light',
			authorName: 'Ada',
			showToolbar: true,
			showFormulaBar: true,
			hiddenActions: ['bold'],
			themeColors: { primary: '#1f9d63' },
		});
		const binding = mountEditor(newHost(), props());
		const element = binding.element;
		// What Editing / Viewing, File > Options, View > Formula Bar and Customize Ribbon do.
		element.readOnly = true;
		element.locale = 'de';
		element.theme = 'dark';
		element.authorName = 'Grace';
		element.showToolbar = false;
		element.showFormulaBar = false;
		element.hiddenActions = ['italic'];
		binding.update({ ...props(), fileName: 'Other.xlsx' });
		expect(element.readOnly).toBe(true);
		expect(element.locale).toBe('de');
		expect(element.theme).toBe('dark');
		expect(element.authorName).toBe('Grace');
		expect(element.showToolbar).toBe(false);
		expect(element.showFormulaBar).toBe(false);
		expect(element.hiddenActions).toEqual(['italic']);
		binding.update({ ...props(), readOnly: false, locale: 'fr', hiddenActions: ['bold', 'x'] });
		expect(element.readOnly).toBe(true);
		expect(element.locale).toBe('fr');
		expect(element.hiddenActions).toEqual(['bold', 'x']);
		binding.update({ ...props(), readOnly: true });
		binding.update({ ...props(), readOnly: false });
		expect(element.readOnly).toBe(false);
		binding.destroy();
	});

	it('compares themeColors by value, so an inline object is applied once', () => {
		const setter = vi.spyOn(XlsxEditorElement.prototype, 'themeColors', 'set');
		const binding = mountEditor(newHost(), { themeColors: { primary: '#123456' } });
		expect(setter).toHaveBeenCalledTimes(1);
		for (let i = 0; i < 3; i++) binding.update({ themeColors: { primary: '#123456' } });
		expect(setter).toHaveBeenCalledTimes(1);
		binding.update({ themeColors: { primary: '#654321' } });
		expect(setter).toHaveBeenCalledTimes(2);
		expect(binding.element.themeColors).toEqual({ primary: '#654321' });
		binding.update({});
		expect(setter).toHaveBeenCalledTimes(3);
		expect(sameThemeColors(undefined, {})).toBe(true);
		expect(sameThemeColors({ primary: '#1' }, { primary: '#1', accent: '#2' })).toBe(false);
		binding.destroy();
	});

	it('reports the user’s Editing / Viewing and Customize Ribbon changes', () => {
		const onReadOnlyChange = vi.fn();
		const onRibbonCustomize = vi.fn();
		const binding = mountEditor(newHost(), { onReadOnlyChange, onRibbonCustomize });
		emit(binding.element, 'readonly-change', { readOnly: true });
		emit(binding.element, 'ribbon-customize', { hiddenActions: ['bold'] });
		expect(onReadOnlyChange).toHaveBeenCalledWith(true);
		expect(onRibbonCustomize).toHaveBeenCalledWith(['bold']);
		binding.destroy();
	});

	it('hears the real element when the user switches to Viewing', () => {
		const onReadOnlyChange = vi.fn();
		const binding = mountEditor(newHost(), { onReadOnlyChange });
		binding.element.readOnly = true;
		expect(onReadOnlyChange).toHaveBeenLastCalledWith(true);
		binding.destroy();
	});
});

describe('src fetching', () => {
	const deferredFetch = () => {
		const pending: ((response: Response) => void)[] = [];
		const fetchMock = vi.fn(
			(_url: string, _init?: RequestInit) =>
				new Promise<Response>((resolve) => pending.push(resolve)),
		);
		vi.stubGlobal('fetch', fetchMock);
		const respond = (index: number) => pending[index]?.(new Response(new Uint8Array([1, 2])));
		return { fetchMock, respond };
	};

	it('ignores a fetch that bytes, a workbook, a cleared or a changed src superseded', async () => {
		const load = vi.spyOn(XlsxEditorElement.prototype, 'load').mockResolvedValue(undefined);
		const { fetchMock, respond } = deferredFetch();
		const binding = mountEditor(newHost(), { src: '/a.xlsx' });
		await flush();
		const bytes = new Uint8Array([9]);
		binding.update({ src: '/a.xlsx', bytes });
		expect(load).toHaveBeenCalledTimes(1);
		expect(fetchMock.mock.calls[0]?.[1]?.signal?.aborted).toBe(true);
		respond(0);
		await flush();
		expect(load).toHaveBeenCalledTimes(1);

		binding.update({ src: '/b.xlsx', bytes });
		await flush();
		const workbook = createWorkbook();
		binding.update({ src: '/b.xlsx', bytes, workbook });
		respond(1);
		await flush();
		expect(load).toHaveBeenCalledTimes(1);
		expect(binding.element.workbook).toBe(workbook);

		binding.update({ src: '/c.xlsx', bytes, workbook });
		await flush();
		binding.update({ bytes, workbook });
		respond(2);
		await flush();
		expect(load).toHaveBeenCalledTimes(1);

		binding.update({ src: '/d.xlsx', bytes, workbook });
		await flush();
		binding.update({ src: '/e.xlsx', bytes, workbook });
		await flush();
		respond(3);
		respond(4);
		await flush();
		await flush();
		expect(load).toHaveBeenCalledTimes(2);
		expect(load).toHaveBeenLastCalledWith(new Uint8Array([1, 2]), 'e.xlsx');
		binding.destroy();
	});

	it('routes every failure in the chain to onWorkbookError, never an unhandled rejection', async () => {
		vi.spyOn(XlsxEditorElement.prototype, 'load').mockResolvedValue(undefined);
		const unhandled = vi.fn();
		process.on('unhandledRejection', unhandled);
		const onWorkbookError = vi.fn();
		vi.stubGlobal('fetch', () => {
			throw new TypeError('fetch is not available');
		});
		const binding = mountEditor(newHost(), { src: '/x.xlsx', onWorkbookError });
		await flush();
		expect(String(onWorkbookError.mock.calls[0]?.[0])).toContain('fetch is not available');
		vi.stubGlobal('fetch', async () => ({
			ok: true,
			arrayBuffer: () => Promise.reject(new Error('body stream broke')),
		}));
		binding.update({ src: '/y.xlsx', onWorkbookError });
		await flush();
		await flush();
		expect(String(onWorkbookError.mock.calls[1]?.[0])).toContain('body stream broke');
		process.off('unhandledRejection', unhandled);
		expect(unhandled).not.toHaveBeenCalled();
		binding.destroy();
	});

	it('names a file whose URL is not valid percent-encoding by its raw segment', async () => {
		const load = vi.spyOn(XlsxEditorElement.prototype, 'load').mockResolvedValue(undefined);
		vi.stubGlobal('fetch', async () => new Response(new Uint8Array([1])));
		const onWorkbookError = vi.fn();
		const binding = mountEditor(newHost(), { src: '/files/100%25%zz.xlsx', onWorkbookError });
		await flush();
		await flush();
		expect(onWorkbookError).not.toHaveBeenCalled();
		expect(load).toHaveBeenCalledWith(new Uint8Array([1]), '100%25%zz.xlsx');
		expect(nameFromUrl('/a/Q3%20Sales.xlsx#top')).toBe('Q3 Sales.xlsx');
		binding.destroy();
	});
});
