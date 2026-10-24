// @vitest-environment jsdom
//
// The print body is sanitised through DOMPurify (see `buildPrintHtmlDocument`
// in `pptx-viewer-shared`), which walks/rewrites the parsed DOM tree.
// happy-dom (this package's default test environment) has a tree-walking bug
// that drops the first of several sibling elements during that rewrite;
// jsdom does not, and is what actually approximates the real browsers this
// code runs in (`window.open` print windows), so this file opts into jsdom
// specifically to get a faithful sanitisation result.
import { mount } from '@vue/test-utils';
import type { PptxHandoutMaster, PptxSlide } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';
import { defineComponent, h, ref } from 'vue';
import { createI18n } from 'vue-i18n';

import { translationsDe } from '../../../../locales/src';
import { toVueI18nSyntax, translationsEn } from '../../i18n';
import type { PrintSettings } from '../components/print-dialog-types';
import { usePrint } from './usePrint';
import type { UsePrintOptions, UsePrintResult } from './usePrint';

function fakeCanvas(id = 'AAAA'): HTMLCanvasElement {
	return { toDataURL: () => `data:image/png;base64,${id}` } as unknown as HTMLCanvasElement;
}

function makeSlides(n: number, notes?: (i: number) => string): PptxSlide[] {
	return Array.from(
		{ length: n },
		(_, i) =>
			({
				id: `s${i}`,
				elements: [
					{
						id: `e${i}`,
						type: 'text',
						text: `Title ${i + 1}`,
						x: 20,
						y: 20,
						width: 300,
						height: 50,
					},
				],
				notes: notes?.(i),
			}) as unknown as PptxSlide,
	);
}

/** `usePrint` set up inside a component of a German viewer (vue-i18n installed). */
function usePrintInGerman(options: UsePrintOptions): UsePrintResult {
	let result: UsePrintResult | undefined;
	const i18n = createI18n({
		legacy: false,
		locale: 'de',
		fallbackLocale: 'en',
		messages: { en: translationsEn, de: toVueI18nSyntax(translationsDe) },
	});
	mount(
		defineComponent({
			setup() {
				// oxlint-disable-next-line react/hooks, react-hooks/rules-of-hooks -- Vue composable, not a React hook
				result = usePrint(options);
				return () => h('div');
			},
		}),
		{ global: { plugins: [i18n] } },
	);
	if (!result) {
		throw new Error('usePrint did not set up');
	}
	return result;
}

function baseSettings(overrides: Partial<PrintSettings> = {}): PrintSettings {
	return {
		printWhat: 'slides',
		orientation: 'landscape',
		colorMode: 'color',
		frameSlides: false,
		slidesPerPage: 6,
		slideRange: 'all',
		customRangeFrom: 1,
		customRangeTo: 1,
		...overrides,
	};
}

describe('usePrint', () => {
	it('opens and closes the dialog', () => {
		const { isPrintDialogOpen, openPrintDialog, closePrintDialog } = usePrint({
			slides: ref(makeSlides(2)),
			activeSlideIndex: ref(0),
			rasterizeSlide: vi.fn(),
		});
		expect(isPrintDialogOpen.value).toBeFalsy();
		openPrintDialog();
		expect(isPrintDialogOpen.value).toBeTruthy();
		closePrintDialog();
		expect(isPrintDialogOpen.value).toBeFalsy();
	});

	it('closes the dialog when printing starts', async () => {
		const openPrintWindow = vi.fn().mockReturnValue(true);
		const result = usePrint({
			slides: ref(makeSlides(1)),
			activeSlideIndex: ref(0),
			rasterizeSlide: vi.fn().mockResolvedValue(fakeCanvas()),
			openPrintWindow,
		});
		result.openPrintDialog();
		await result.print(baseSettings());
		expect(result.isPrintDialogOpen.value).toBeFalsy();
	});

	it('prints direct slides as SVG without rasterising', async () => {
		const rasterizeSlide = vi.fn().mockResolvedValue(fakeCanvas());
		const openPrintWindow = vi.fn().mockReturnValue(true);
		const { print } = usePrint({
			slides: ref(makeSlides(3)),
			activeSlideIndex: ref(0),
			rasterizeSlide,
			openPrintWindow,
		});
		await print(baseSettings());
		expect(rasterizeSlide).not.toHaveBeenCalled();
		expect(openPrintWindow).toHaveBeenCalledOnce();
		const html = openPrintWindow.mock.calls[0][0] as string;
		expect(html).toContain('<svg');
		expect(html).toContain('Title 1');
		expect(html).toContain('size: landscape;');
	});

	it('only exports the active slide for the "current" range', async () => {
		const rasterizeSlide = vi.fn().mockResolvedValue(fakeCanvas());
		const openPrintWindow = vi.fn().mockReturnValue(true);
		const { print } = usePrint({
			slides: ref(makeSlides(5)),
			activeSlideIndex: ref(2),
			rasterizeSlide,
			openPrintWindow,
		});
		await print(baseSettings({ slideRange: 'current' }));
		expect(rasterizeSlide).not.toHaveBeenCalled();
		const html = openPrintWindow.mock.calls[0][0] as string;
		expect(html).toContain('Title 3');
		expect(html).not.toContain('Title 2');
	});

	it('honours a custom range', async () => {
		const rasterizeSlide = vi.fn().mockResolvedValue(fakeCanvas());
		const openPrintWindow = vi.fn().mockReturnValue(true);
		const { print } = usePrint({
			slides: ref(makeSlides(10)),
			activeSlideIndex: ref(0),
			rasterizeSlide,
			openPrintWindow,
		});
		await print(baseSettings({ slideRange: 'custom', customRangeFrom: 3, customRangeTo: 5 }));
		expect(rasterizeSlide).not.toHaveBeenCalled();
		const html = openPrintWindow.mock.calls[0][0] as string;
		expect(html).toContain('Title 3');
		expect(html).toContain('Title 5');
		expect(html).not.toContain('Title 2');
		expect(html).not.toContain('Title 6');
	});

	it('prints chart geometry through the core SVG exporter', async () => {
		const chartSlide: PptxSlide = {
			id: 'chart-slide',
			rId: 'rId1',
			slideNumber: 1,
			elements: [
				{
					type: 'chart',
					id: 'chart1',
					x: 20,
					y: 20,
					width: 400,
					height: 240,
					chartData: {
						chartType: 'bar',
						categories: ['A', 'B'],
						series: [{ name: 'Values', values: [10, 20], color: '#123456' }],
					},
				},
			],
		};
		const rasterizeSlide = vi.fn();
		const openPrintWindow = vi.fn().mockReturnValue(true);
		const { print } = usePrint({
			slides: ref([chartSlide]),
			activeSlideIndex: ref(0),
			rasterizeSlide,
			openPrintWindow,
		});

		await print(baseSettings());

		expect(rasterizeSlide).not.toHaveBeenCalled();
		const html = openPrintWindow.mock.calls[0][0] as string;
		expect(html).toContain('data-pptx-element="chart"');
		expect(html).toContain('data-chart-mark="bar"');
		expect(html).not.toContain('>chart</text>');
	});

	it('prints the outline without rasterising', async () => {
		const rasterizeSlide = vi.fn();
		const openPrintWindow = vi.fn().mockReturnValue(true);
		const { print } = usePrint({
			slides: ref(makeSlides(2, (i) => (i === 0 ? 'note one' : ''))),
			activeSlideIndex: ref(0),
			rasterizeSlide,
			openPrintWindow,
		});
		await print(baseSettings({ printWhat: 'outline' }));
		expect(rasterizeSlide).not.toHaveBeenCalled();
		const html = openPrintWindow.mock.calls[0][0] as string;
		expect(html).toContain('<h2>Title 1</h2>');
		expect(html).toContain('<p>note one</p>');
		expect(html).toContain('outline-page');
	});

	it('emits notes pages with escaped notes text', async () => {
		const rasterizeSlide = vi.fn().mockResolvedValue(fakeCanvas());
		const openPrintWindow = vi.fn().mockReturnValue(true);
		const { print } = usePrint({
			slides: ref(makeSlides(1, () => 'a & <b>')),
			activeSlideIndex: ref(0),
			rasterizeSlide,
			openPrintWindow,
		});
		await print(baseSettings({ printWhat: 'notes' }));
		const html = openPrintWindow.mock.calls[0][0] as string;
		expect(html).toContain('notes-page');
		expect(html).toContain('a &amp; &lt;b&gt;');
		// notes are always portrait
		expect(html).toContain('@page { size: portrait;');
	});

	it('paginates handouts by slides-per-page', async () => {
		const rasterizeSlide = vi.fn().mockResolvedValue(fakeCanvas());
		const openPrintWindow = vi.fn().mockReturnValue(true);
		const { print } = usePrint({
			slides: ref(makeSlides(7)),
			activeSlideIndex: ref(0),
			rasterizeSlide,
			openPrintWindow,
		});
		await print(baseSettings({ printWhat: 'handouts', slidesPerPage: 6 }));
		const html = openPrintWindow.mock.calls[0][0] as string;
		// 7 slides / 6 per page → 2 pages
		expect(html.match(/class="page"/gu)).toHaveLength(2);
		expect(html).toContain('handout-grid');
	});

	it('paints the handout master footer text when a handoutMaster ref is wired and ftr is enabled', async () => {
		const rasterizeSlide = vi.fn().mockResolvedValue(fakeCanvas());
		const openPrintWindow = vi.fn().mockReturnValue(true);
		const handoutMaster: PptxHandoutMaster = {
			path: 'ppt/handoutMasters/handoutMaster1.xml',
			slidesPerPage: 6,
			headerFooter: { hasFooter: true },
			elements: [
				{
					id: 'ftr1',
					type: 'text',
					placeholderType: 'ftr',
					x: 0,
					y: 0,
					width: 100,
					height: 20,
					text: 'Confidential - Acme Corp',
				} as unknown as PptxSlide['elements'][number],
			],
		};
		const { print } = usePrint({
			slides: ref(makeSlides(2)),
			activeSlideIndex: ref(0),
			rasterizeSlide,
			openPrintWindow,
			handoutMaster: ref(handoutMaster),
		});
		await print(baseSettings({ printWhat: 'handouts', slidesPerPage: 6 }));
		const html = openPrintWindow.mock.calls[0][0] as string;
		expect(html).toContain('Confidential - Acme Corp');
	});

	it('uses the 3-per-page note-line layout', async () => {
		const rasterizeSlide = vi.fn().mockResolvedValue(fakeCanvas());
		const openPrintWindow = vi.fn().mockReturnValue(true);
		const { print } = usePrint({
			slides: ref(makeSlides(3)),
			activeSlideIndex: ref(0),
			rasterizeSlide,
			openPrintWindow,
		});
		await print(baseSettings({ printWhat: 'handouts', slidesPerPage: 3 }));
		const html = openPrintWindow.mock.calls[0][0] as string;
		expect(html).toContain('handout-grid-3');
		expect(html).toContain('handout-note-line');
	});

	it('applies the grayscale colour filter', async () => {
		const openPrintWindow = vi.fn().mockReturnValue(true);
		const { print } = usePrint({
			slides: ref(makeSlides(1)),
			activeSlideIndex: ref(0),
			rasterizeSlide: vi.fn().mockResolvedValue(fakeCanvas()),
			openPrintWindow,
		});
		await print(baseSettings({ colorMode: 'grayscale' }));
		expect(openPrintWindow.mock.calls[0][0]).toContain('filter: grayscale(1);');
	});

	it('does nothing when no slides are selected (empty deck)', async () => {
		const rasterizeSlide = vi.fn();
		const openPrintWindow = vi.fn().mockReturnValue(true);
		const { print } = usePrint({
			slides: ref(makeSlides(0)),
			activeSlideIndex: ref(0),
			rasterizeSlide,
			openPrintWindow,
		});
		await print(baseSettings());
		expect(rasterizeSlide).not.toHaveBeenCalled();
		expect(openPrintWindow).not.toHaveBeenCalled();
	});

	it('opens nothing for an outline of hidden slides only', async () => {
		const openPrintWindow = vi.fn().mockReturnValue(true);
		const slideList = makeSlides(2).map((slide) => ({ ...slide, hidden: true }));
		const { print } = usePrint({
			slides: ref(slideList),
			activeSlideIndex: ref(0),
			rasterizeSlide: vi.fn(),
			openPrintWindow,
		});
		await print(baseSettings({ printWhat: 'outline' }));
		expect(openPrintWindow).not.toHaveBeenCalled();
	});

	it('titles the print window in the viewer language', async () => {
		const openPrintWindow = vi.fn().mockReturnValue(true);
		const { print } = usePrintInGerman({
			slides: ref(makeSlides(2)),
			activeSlideIndex: ref(0),
			rasterizeSlide: vi.fn().mockResolvedValue(fakeCanvas()),
			openPrintWindow,
		});
		const titles: string[] = [];
		for (const settings of [
			baseSettings(),
			baseSettings({ printWhat: 'outline' }),
			baseSettings({ printWhat: 'notes' }),
			baseSettings({ printWhat: 'handouts', slidesPerPage: 4 }),
		]) {
			await print(settings);
			titles.push(
				/<title>(?<title>[^<]*)<\/title>/u.exec(openPrintWindow.mock.calls.at(-1)?.[0])?.groups
					?.title ?? '',
			);
		}
		expect(titles).toStrictEqual([
			'Folien',
			'Gliederung',
			'Notizseiten',
			'4 Folien pro Handzettel',
		]);
	});

	it('captures the slides when the vector export throws', async () => {
		const rasterizeSlide = vi.fn().mockResolvedValue(fakeCanvas());
		const openPrintWindow = vi.fn().mockReturnValue(true);
		const warn = vi.spyOn(console, 'warn').mockImplementation(() => {});
		const broken = [{ ...makeSlides(1)[0], elements: undefined }] as unknown as PptxSlide[];
		const { print } = usePrint({
			slides: ref(broken),
			activeSlideIndex: ref(0),
			rasterizeSlide,
			openPrintWindow,
		});
		await print(baseSettings());
		expect(rasterizeSlide).toHaveBeenCalledExactlyOnceWith(0, 1);
		expect(openPrintWindow.mock.calls[0][0]).toContain('data:image/png;base64,AAAA');
		warn.mockRestore();
	});

	it('swallows rasterisation errors', async () => {
		const openPrintWindow = vi.fn().mockReturnValue(true);
		const errSpy = vi.spyOn(console, 'error').mockImplementation(() => {});
		const { print } = usePrint({
			slides: ref(makeSlides(2)),
			activeSlideIndex: ref(0),
			rasterizeSlide: vi.fn().mockRejectedValue(new Error('boom')),
			openPrintWindow,
		});
		await expect(print(baseSettings({ printWhat: 'notes' }))).resolves.toBeUndefined();
		expect(openPrintWindow).not.toHaveBeenCalled();
		errSpy.mockRestore();
	});
});
