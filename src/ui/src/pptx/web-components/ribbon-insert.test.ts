// @vitest-environment jsdom
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';

import { INSERT_GROUPS } from '../render';
import type { RibbonInsertState } from '../render';
import { registerPptxWebControls } from './index';

beforeAll(registerPptxWebControls);
afterEach(() => document.body.replaceChildren());

const STATE: RibbonInsertState = {
	editable: true,
	hasSelection: true,
	shapeType: 'ellipse',
	chartKind: 'pie',
	activeFreeformTool: 'curve',
	freeformTools: ['freeformShape', 'curve'],
	newSlideAvailable: true,
	commentAvailable: true,
	layouts: [
		{ path: 'ppt/slideLayouts/slideLayout1.xml', name: 'Title Slide' },
		{ path: 'ppt/slideLayouts/slideLayout2.xml', name: 'Title and Content' },
	],
	currentLayoutPath: 'ppt/slideLayouts/slideLayout2.xml',
};
/** PowerPoint's Insert tab, left to right. */
const GROUPS = [
	'insert.slides',
	'insert.tables',
	'insert.images',
	'insert.camera',
	'insert.illustrations',
	'insert.forms',
	'insert.addIns',
	'insert.links',
	'insert.comments',
	'insert.text',
	'insert.symbols',
	'insert.media',
];
const CAPTIONS = [
	'Slides',
	'Tables',
	'Images',
	'Camera',
	'Illustrations',
	'Forms',
	'Add-ins',
	'Links',
	'Comments',
	'Text',
	'Symbols',
	'Media',
];
const CONTROLS = [
	'insert.slides.newSlide',
	'insert.slides.reuseSlides',
	'insert.tables.table',
	'insert.images.pictures',
	'insert.images.screenshot',
	'insert.images.photoAlbum',
	'insert.camera.cameo',
	'insert.illustrations.shapes',
	'insert.illustrations.icons',
	'insert.illustrations.threeDModels',
	'insert.illustrations.smartArt',
	'insert.illustrations.chart',
	'insert.forms.forms',
	'insert.addIns.getAddIns',
	'insert.addIns.myAddIns',
	'insert.links.zoom',
	'insert.links.link',
	'insert.links.action',
	'insert.comments.comment',
	'insert.text.textBox',
	'insert.text.headerFooter',
	'insert.text.wordArt',
	'insert.text.dateTime',
	'insert.text.slideNumber',
	'insert.text.object',
	'insert.symbols.equation',
	'insert.symbols.symbol',
	'insert.media.media',
	'insert.media.audio',
	'insert.media.screenRecording',
];

function mount(state: Partial<RibbonInsertState> = {}) {
	const host = document.createElement('pptx-ui-ribbon-insert');
	host.state = { ...STATE, ...state };
	document.body.append(host);
	const requests = vi.fn();
	host.addEventListener('insert-request', (event) => requests((event as CustomEvent).detail));
	return { host, requests };
}
const control = (host: HTMLElement, id: string) =>
	host.querySelector<HTMLElement>(`[data-ribbon-control="${id}"]`)!;
/** The inner button of a plain command, or the opening button of a menu. */
const button = (host: HTMLElement, id: string) => {
	const el = control(host, id);
	return (el.shadowRoot?.querySelector('button') ??
		el.querySelector('.trigger')) as HTMLButtonElement;
};
const calls = (requests: ReturnType<typeof vi.fn>) => requests.mock.calls.map(([detail]) => detail);

describe('shared Insert view', () => {
	it('lists the groups in PowerPoint order with its captions', () => {
		const { host } = mount();
		const groups = [...host.querySelectorAll<HTMLElement>('[data-ribbon-group]')];
		expect(groups.map((el) => el.dataset.ribbonGroup)).toStrictEqual(GROUPS);
		expect(groups.map((el) => el.getAttribute('label'))).toStrictEqual(CAPTIONS);
		expect(INSERT_GROUPS.map((group) => group.id)).toStrictEqual(GROUPS);
	});

	it('exposes every public group and control id exactly once', () => {
		const { host } = mount();
		for (const id of GROUPS) {
			expect(host.querySelectorAll(`[data-ribbon-group="${id}"]`)).toHaveLength(1);
		}
		for (const id of CONTROLS) {
			expect(host.querySelectorAll(`[data-ribbon-control="${id}"]`)).toHaveLength(1);
		}
	});

	it('dispatches one typed intent per command and keeps Link selection-gated', () => {
		const { host, requests } = mount();
		button(host, 'insert.text.textBox').click();
		button(host, 'insert.comments.comment').click();
		button(host, 'insert.text.dateTime').click();
		button(host, 'insert.text.slideNumber').click();
		expect(calls(requests)).toStrictEqual([
			{ kind: 'command', value: 'textBox' },
			{ kind: 'command', value: 'comment' },
			{ kind: 'field', value: 'datetime' },
			{ kind: 'field', value: 'slidenum' },
		]);
		host.state = { ...STATE, hasSelection: false };
		expect(button(host, 'insert.links.link').disabled).toBeTruthy();
		expect(button(host, 'insert.text.textBox').disabled).toBeFalsy();
		button(host, 'insert.links.link').click();
		expect(requests).toHaveBeenCalledTimes(4);
	});

	it('opens the Table, Pictures, Video and Audio menus and picks their working entries', () => {
		const { host, requests } = mount();
		const pick = (id: string, item: string) => {
			button(host, id).click();
			control(host, id).querySelector<HTMLButtonElement>(`[data-insert-item="${item}"]`)!.click();
		};
		pick('insert.tables.table', 'insertTable');
		pick('insert.images.pictures', 'thisDevice');
		pick('insert.media.media', 'videoThisDevice');
		pick('insert.media.audio', 'audioOnMyPc');
		pick('insert.symbols.equation', 'newEquation');
		expect(calls(requests)).toStrictEqual([
			{ kind: 'command', value: 'table' },
			{ kind: 'command', value: 'image' },
			{ kind: 'command', value: 'media' },
			{ kind: 'command', value: 'media' },
			{ kind: 'command', value: 'equation' },
		]);
		const table = control(host, 'insert.tables.table');
		button(host, 'insert.tables.table').click();
		const draw = table.querySelector<HTMLButtonElement>('[data-insert-item="drawTable"]')!;
		expect(draw.disabled).toBeTruthy();
		expect(draw.title).toBe('Not available in this viewer yet.');
		draw.click();
		expect(requests).toHaveBeenCalledTimes(5);
	});

	it('sizes a table on the Table grid: hover shows the size, a click inserts it', () => {
		const { host, requests } = mount();
		const table = control(host, 'insert.tables.table');
		button(host, 'insert.tables.table').click();
		const cells = table.querySelectorAll('.table-picker-cell');
		expect(cells).toHaveLength(80);
		const label = table.querySelector('.table-picker-label')!;
		expect(label.textContent).toBe('Insert Table');
		const cell = table.querySelector<HTMLElement>('[data-row="4"][data-column="3"]')!;
		cell.dispatchEvent(new Event('pointerover', { bubbles: true }));
		expect(label.textContent).toBe('3x4 Table');
		expect(table.querySelectorAll('.table-picker-cell[data-on]')).toHaveLength(12);
		cell.click();
		expect(calls(requests)).toStrictEqual([{ kind: 'table', rows: 4, columns: 3 }]);
		expect(button(host, 'insert.tables.table').getAttribute('aria-expanded')).toBe('false');
	});

	it('sizes a table from the keyboard and refuses it when read-only', () => {
		const { host, requests } = mount();
		const table = control(host, 'insert.tables.table');
		button(host, 'insert.tables.table').click();
		const grid = table.querySelector<HTMLElement>('.table-picker-grid')!;
		grid.focus();
		for (const key of ['ArrowRight', 'ArrowRight', 'ArrowDown', 'Enter']) {
			grid.dispatchEvent(new KeyboardEvent('keydown', { key, bubbles: true }));
		}
		expect(calls(requests)).toStrictEqual([{ kind: 'table', rows: 2, columns: 3 }]);
		host.state = { ...STATE, editable: false };
		expect(grid.hasAttribute('aria-disabled')).toBeTruthy();
	});

	it('runs Equation from its face and keeps New Slide a split button over the layouts', () => {
		const { host, requests } = mount();
		const equation = control(host, 'insert.symbols.equation');
		equation.querySelector<HTMLButtonElement>('.primary')!.click();
		const slide = control(host, 'insert.slides.newSlide');
		slide.querySelector<HTMLButtonElement>('.primary')!.click();
		expect(calls(requests)).toStrictEqual([
			{ kind: 'command', value: 'equation' },
			{ kind: 'newSlide', value: null },
		]);
		const caret = slide.querySelector<HTMLButtonElement>('.trigger')!;
		caret.click();
		const items = [...slide.querySelectorAll<HTMLButtonElement>('[data-insert-item]')];
		expect(items.map((item) => item.textContent)).toStrictEqual([
			'Title Slide',
			'Title and Content',
		]);
		expect(items[1].getAttribute('aria-current')).toBe('true');
		items[0].click();
		expect(calls(requests).at(-1)).toStrictEqual({
			kind: 'newSlide',
			value: 'ppt/slideLayouts/slideLayout1.xml',
		});
		host.state = { ...STATE, newSlideAvailable: false };
		expect(slide.querySelector<HTMLButtonElement>('.primary')!.disabled).toBeTruthy();
		expect(slide.querySelector<HTMLButtonElement>('.trigger')!.disabled).toBeTruthy();
	});

	it('ends the New Slide menu with Slide Templates, which a deck with no layouts still reaches', () => {
		const { host, requests } = mount({ slideTemplatesAvailable: true });
		const slide = control(host, 'insert.slides.newSlide');
		slide.querySelector<HTMLButtonElement>('.trigger')!.click();
		const items = [...slide.querySelectorAll<HTMLButtonElement>('[data-insert-item]')];
		expect(items.at(-1)!.textContent).toBe('Slide Templates...');
		items.at(-1)!.click();
		expect(calls(requests).at(-1)).toStrictEqual({ kind: 'slideTemplates' });
		// No layouts and nothing to insert from them: the face is off, the menu is not.
		host.state = {
			...STATE,
			layouts: [],
			newSlideAvailable: false,
			slideTemplatesAvailable: true,
		};
		expect(slide.querySelector<HTMLButtonElement>('.primary')!.disabled).toBeTruthy();
		const caret = slide.querySelector<HTMLButtonElement>('.trigger')!;
		expect(caret.disabled).toBeFalsy();
		caret.click();
		expect(
			[...slide.querySelectorAll<HTMLButtonElement>('[data-insert-item]')].map(
				(item) => item.textContent,
			),
		).toStrictEqual(['Slide Templates...']);
		// Without the capability the menu is empty again and the caret goes.
		host.state = { ...STATE, layouts: [], newSlideAvailable: false };
		expect(slide.querySelector<HTMLButtonElement>('.trigger')!.disabled).toBeTruthy();
	});

	it('shows what the editor cannot do disabled, with an honest ScreenTip', () => {
		const { host, requests } = mount();
		for (const id of [
			'insert.slides.reuseSlides',
			'insert.images.screenshot',
			'insert.images.photoAlbum',
			'insert.camera.cameo',
			'insert.illustrations.icons',
			'insert.illustrations.threeDModels',
			'insert.forms.forms',
			'insert.addIns.getAddIns',
			'insert.addIns.myAddIns',
			'insert.links.zoom',
			'insert.text.wordArt',
			'insert.text.object',
			'insert.symbols.symbol',
			'insert.media.screenRecording',
		]) {
			const inner = button(host, id);
			expect(inner.disabled, id).toBeTruthy();
			const tip = (control(host, id).getAttribute('title') ?? inner.title).split('\n');
			expect(tip, id).toHaveLength(3);
			expect(tip[2], id).toBe('Not available in this viewer yet.');
			inner.click();
		}
		expect(requests).not.toHaveBeenCalled();
		const reuse = control(host, 'insert.slides.reuseSlides').getAttribute('title')!.split('\n');
		expect(reuse[0]).toBe('Reuse Slides');
		expect(reuse[1]).toBe('Insert slides from another presentation.');
	});

	it('disables Comment and New Slide when the host does not offer them', () => {
		const { host, requests } = mount({ commentAvailable: false, newSlideAvailable: false });
		expect(button(host, 'insert.comments.comment').disabled).toBeTruthy();
		expect(control(host, 'insert.comments.comment').getAttribute('title')).toContain(
			'Not available in this viewer yet.',
		);
		button(host, 'insert.comments.comment').click();
		expect(requests).not.toHaveBeenCalled();
	});

	it('is read-only when not editable and rejects direct requests', () => {
		const { host, requests } = mount({ editable: false });
		expect(button(host, 'insert.tables.table').disabled).toBeTruthy();
		button(host, 'insert.tables.table').click();
		host.querySelector<HTMLElement>('[data-pptx-drawing-tool="curve"]')!.click();
		for (const id of ['insert.illustrations.shapes', 'insert.illustrations.chart']) {
			expect(button(host, id).disabled).toBeTruthy();
		}
		expect(requests).not.toHaveBeenCalled();
	});

	it('stages a gallery pick and then inserts it, as Office does', () => {
		vi.useFakeTimers();
		const { host, requests } = mount();
		const pick = (id: string, value: string) => {
			const menu = control(host, id);
			menu.querySelector<HTMLButtonElement>('.trigger')!.click();
			menu.querySelector<HTMLButtonElement>(`[data-insert-item="${value}"]`)!.click();
		};
		pick('insert.illustrations.shapes', 'star5');
		pick('insert.illustrations.chart', 'pie');
		vi.runAllTimers();
		vi.useRealTimers();
		expect(calls(requests)).toStrictEqual([
			{ kind: 'shapeType', value: 'star5' },
			{ kind: 'chartType', value: 'pie' },
			{ kind: 'shape', value: 'star5' },
			{ kind: 'chart', value: 'pie' },
		]);
	});

	it('lays the controls out as large or small Office commands', () => {
		const { host } = mount();
		for (const id of ['insert.text.textBox', 'insert.comments.comment', 'insert.links.link']) {
			expect(control(host, id).hasAttribute('compact'), id).toBeFalsy();
		}
		for (const id of ['insert.text.dateTime', 'insert.text.slideNumber', 'insert.text.object']) {
			expect(control(host, id).hasAttribute('compact'), id).toBeTruthy();
		}
		for (const id of [
			'insert.illustrations.shapes',
			'insert.links.action',
			'insert.tables.table',
			'insert.media.audio',
		]) {
			expect(control(host, id).querySelector('.trigger.large'), id).not.toBeNull();
		}
		expect(control(host, 'insert.illustrations.shapes').querySelector('.grid')).not.toBeNull();
		const stacks = host.querySelectorAll('[data-ribbon-group="insert.text"] > .stack');
		expect(stacks).toHaveLength(1);
		expect(stacks[0].children).toHaveLength(3);
	});

	it('reflects the armed Freeform tool and toggles it off on a second press', () => {
		const { host, requests } = mount();
		const curve = host.querySelector<HTMLElement>('[data-pptx-drawing-tool="curve"]')!;
		const inner = curve.shadowRoot!.querySelector('button')!;
		expect(inner.getAttribute('aria-pressed')).toBe('true');
		curve.click();
		host.querySelector<HTMLElement>('[data-pptx-drawing-tool="freeformShape"]')!.click();
		expect(calls(requests)).toStrictEqual([
			{ kind: 'freeform', value: null },
			{ kind: 'freeform', value: 'freeformShape' },
		]);
		host.state = { ...STATE, freeformTools: ['curve'] };
		expect(
			host.querySelector<HTMLElement>('[data-pptx-drawing-tool="freeformShape"]')!.hidden,
		).toBeTruthy();
	});

	it('opens menus from the keyboard, picks items, closes on Escape and outside press', () => {
		const { host, requests } = mount();
		const action = control(host, 'insert.links.action');
		const trigger = action.querySelector<HTMLButtonElement>('.trigger')!;
		trigger.focus();
		trigger.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
		expect(trigger.getAttribute('aria-expanded')).toBe('true');
		const items = action.querySelectorAll<HTMLButtonElement>('[role=menuitem]');
		expect(items.length).toBeGreaterThan(5);
		expect(document.activeElement).toBe(items[0]);
		document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
		expect(trigger.getAttribute('aria-expanded')).toBe('false');
		expect(document.activeElement).toBe(trigger);
		trigger.click();
		items[1].click();
		expect(requests.mock.calls[0][0].kind).toBe('actionButton');
		expect(trigger.getAttribute('aria-expanded')).toBe('false');
		const table = button(host, 'insert.tables.table');
		table.click();
		expect(table.getAttribute('aria-expanded')).toBe('true');
		document.body.dispatchEvent(new Event('pointerdown', { bubbles: true }));
		expect(table.getAttribute('aria-expanded')).toBe('false');
	});

	it('skips disabled entries when arrowing through a menu', () => {
		const { host } = mount();
		const menu = control(host, 'insert.images.pictures');
		const trigger = menu.querySelector<HTMLButtonElement>('.trigger')!;
		trigger.focus();
		trigger.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true }));
		const enabled = [...menu.querySelectorAll<HTMLButtonElement>('[role=menuitem]:not(:disabled)')];
		expect(enabled.map((item) => item.dataset.insertItem)).toStrictEqual(['thisDevice']);
		expect(document.activeElement).toBe(enabled[0]);
	});

	it('hides unavailable capabilities and keeps independent instances and focus', () => {
		const first = mount();
		const second = mount({
			chartAvailable: false,
			fieldAvailable: false,
			headerFooterAvailable: false,
		});
		for (const id of [
			'insert.illustrations.chart',
			'insert.text.dateTime',
			'insert.text.slideNumber',
			'insert.text.headerFooter',
		]) {
			expect(control(second.host, id).hidden, id).toBeTruthy();
			expect(control(first.host, id).hidden, id).toBeFalsy();
		}
		const table = button(first.host, 'insert.tables.table');
		table.focus();
		first.host.state = { ...STATE, shapeType: 'star5' };
		expect(button(first.host, 'insert.tables.table')).toBe(table);
		first.host.remove();
		document.body.append(first.host);
		expect(first.host.querySelectorAll('[data-ribbon-control="insert.tables.table"]')).toHaveLength(
			1,
		);
	});

	it('uses translated labels with English fallbacks', () => {
		const { host } = mount({ translate: (key) => (key === 'pptx.ribbon.table' ? 'Tabelle' : key) });
		expect(control(host, 'insert.tables.table').querySelector('.label')!.textContent).toBe(
			'Tabelle',
		);
		expect(control(host, 'insert.images.pictures').querySelector('.label')!.textContent).toBe(
			'Pictures',
		);
		expect(control(host, 'insert.text.dateTime').getAttribute('label')).toBe('Date & Time');
	});
});
