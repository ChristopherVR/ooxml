import {
	CATEGORIES,
	INSERT_CHART_TYPES,
	PRESETS,
	RECENT_SHAPES_KEY,
	SHAPE_GALLERY_ITEMS,
	SHAPE_PRESET_DEFS,
} from 'ooxml-ui/pptx';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../../../i18n';
import type { RibbonInsertHandlers } from '../ribbon-types';
import { createInsertTab } from './insert-tab';

function makeHandlers(over: Partial<RibbonInsertHandlers> = {}): RibbonInsertHandlers {
	return {
		insert: vi.fn(),
		insertImage: vi.fn(),
		insertMedia: vi.fn(),
		insertChart: vi.fn(),
		insertSmartArt: vi.fn(),
		insertEquation: vi.fn(),
		insertActionButton: vi.fn(),
		insertField: vi.fn(),
		...over,
	};
}

function make(
	over: Partial<RibbonInsertHandlers> = {},
	spies: Record<string, () => void> = {},
	templates?: Parameters<typeof createInsertTab>[7],
) {
	const tab = createInsertTab(
		document,
		createTranslator(),
		makeHandlers(over),
		spies.equation ?? vi.fn(),
		spies.headerFooter ?? vi.fn(),
		spies.hyperlink ?? vi.fn(),
		spies.comments,
		templates,
	);
	document.body.append(tab.el);
	return tab;
}
function control(tab: { el: HTMLElement }, id: string): HTMLButtonElement {
	return tab.el
		.querySelector(`[data-ribbon-control="${id}"]`)!
		.shadowRoot!.querySelector<HTMLButtonElement>('button')!;
}
/** Open a gallery menu (Shapes, Chart) and choose an entry; the pick stages, then inserts. */
function choose(tab: { el: HTMLElement }, id: string, value: string): void {
	tab.el.querySelector<HTMLButtonElement>(`[data-ribbon-control="${id}"] .trigger`)!.click();
	tab.el
		.querySelector<HTMLButtonElement>(
			`[data-ribbon-control="${id}"] [data-insert-item="${value}"]`,
		)!
		.click();
	vi.runAllTimers();
}
/** Open a dropdown and pick one of its entries (no staging timer). */
function pick(tab: { el: HTMLElement }, id: string, value: string): void {
	tab.el.querySelector<HTMLButtonElement>(`[data-ribbon-control="${id}"] .trigger`)!.click();
	tab.el
		.querySelector<HTMLButtonElement>(
			`[data-ribbon-control="${id}"] [data-insert-item="${value}"]`,
		)!
		.click();
}
type Shadowed = HTMLElement & { shadowRoot: ShadowRoot };
/** The shared Insert SmartArt / Insert Chart dialog element, and its picker's shadow root. */
function shared(tag: string): HTMLElement & { open: boolean } {
	return document.querySelector(tag) as HTMLElement & { open: boolean };
}
function pickerRoot(tag: string, picker: string): ShadowRoot {
	return (shared(tag).querySelector(picker) as Shadowed).shadowRoot;
}
function footerButton(tag: string, id: string): HTMLButtonElement {
	return (
		shared(tag).querySelector('office-ui-dialog-footer') as Shadowed
	).shadowRoot.querySelector(`button[data-action="${id}"]`)!;
}
const SMARTART = 'pptx-ui-insert-smartart-dialog';
const CHART = 'pptx-ui-insert-chart-dialog';

describe('createInsertTab', () => {
	afterEach(() => {
		document.body.replaceChildren();
		// A shape pick is remembered in Recently Used Shapes.
		localStorage.removeItem(RECENT_SHAPES_KEY);
	});

	it('ends the New Slide menu with Slide Templates, even for a deck with no layouts', () => {
		const insert = vi.fn();
		const tab = make({}, {}, { insert, available: () => true });
		const newSlide = tab.el.querySelector<HTMLElement>(
			'[data-ribbon-control="insert.slides.newSlide"]',
		)!;
		newSlide.querySelector<HTMLButtonElement>('.trigger')!.click();
		const items = [...newSlide.querySelectorAll<HTMLButtonElement>('[data-insert-item]')];
		expect(items.map((item) => item.textContent)).toStrictEqual(['Slide Templates...']);
		items[0].click();
		expect(document.querySelector('[role="dialog"]')).not.toBeNull();
	});

	it('leaves the entry out when the host has no templates or hides the dialog', () => {
		for (const tab of [make(), make({}, {}, { insert: vi.fn(), available: () => false })]) {
			const trigger = tab.el.querySelector<HTMLButtonElement>(
				'[data-ribbon-control="insert.slides.newSlide"] .trigger',
			)!;
			expect(trigger.disabled).toBeTruthy();
		}
	});

	it('offers the Freeform: Shape and Curve tools and arms / disarms them', () => {
		const armFreeformTool = vi.fn();
		const tab = make({ armFreeformTool, visibleDrawingTools: () => ['freeformShape', 'curve'] });
		const freeform = tab.el.querySelector<HTMLElement>('[data-pptx-drawing-tool="freeformShape"]')!;
		const curve = tab.el.querySelector<HTMLElement>('[data-pptx-drawing-tool="curve"]')!;
		const inner = curve.shadowRoot!.querySelector('button')!;
		expect(freeform.getAttribute('label')).toBe('Freeform: Shape');
		expect(inner.getAttribute('aria-pressed')).toBe('false');
		curve.click();
		expect(armFreeformTool).toHaveBeenLastCalledWith('curve');
		tab.setFreeformTool('curve');
		expect(inner.getAttribute('aria-pressed')).toBe('true');
		curve.click();
		expect(armFreeformTool).toHaveBeenLastCalledWith(null);
	});

	it('leaves the drawing tools out when the host hides them', () => {
		const tab = make({ armFreeformTool: vi.fn(), visibleDrawingTools: () => ['curve'] });
		expect(
			tab.el.querySelector<HTMLElement>('[data-pptx-drawing-tool="freeformShape"]')!.hidden,
		).toBeTruthy();
		expect(
			tab.el.querySelector<HTMLElement>('[data-pptx-drawing-tool="curve"]')!.hidden,
		).toBeFalsy();
		const none = make();
		expect(none.el.querySelector<HTMLElement>('.stack')!.hidden).toBeTruthy();
	});

	it('renders the shared Insert element with every public group and control id', () => {
		const tab = make();
		expect(tab.el.querySelector('pptx-ui-ribbon-insert')).not.toBeNull();
		for (const id of [
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
		]) {
			expect(tab.el.querySelectorAll(`[data-ribbon-group="${id}"]`)).toHaveLength(1);
		}
		const items = (id: string) =>
			tab.el.querySelectorAll(`[data-ribbon-control="${id}"] [data-insert-item]`);
		expect(items('insert.illustrations.shapes')).toHaveLength(SHAPE_GALLERY_ITEMS.length);
		// Every chart type, then Insert Chart... for the whole dialog.
		expect(items('insert.illustrations.chart')).toHaveLength(INSERT_CHART_TYPES.length + 1);
		// Shapes, Chart, Action, Table, Pictures and Video are large dropdowns, not selects beside a button.
		expect(tab.el.querySelectorAll('select')).toHaveLength(0);
		expect(tab.el.querySelectorAll('.trigger.large')).toHaveLength(10);
	});

	it('opens the hyperlink editor and needs a selection to be usable', () => {
		const hyperlink = vi.fn();
		const tab = make({}, { hyperlink });
		const link = control(tab, 'insert.links.link');
		// A link always attaches to something, so an empty selection is enough to
		// rule the command out even on an editable deck.
		expect(link.disabled).toBeTruthy();
		tab.setEditable(true);
		expect(link.disabled).toBeTruthy();
		tab.setHasSelection(true);
		expect(link.disabled).toBeFalsy();
		link.click();
		expect(hyperlink).toHaveBeenCalledOnce();
		tab.setHasSelection(false);
		expect(link.disabled).toBeTruthy();
	});

	it('dispatches insert("text") / insert("table") for the fixed buttons', () => {
		const insert = vi.fn();
		const tab = make({ insert });
		control(tab, 'insert.text.textBox').click();
		pick(tab, 'insert.tables.table', 'insertTable');
		expect(insert).toHaveBeenCalledWith('text');
		expect(insert).toHaveBeenCalledWith('table');
	});

	it('inserts the shape chosen from the Shapes gallery', () => {
		vi.useFakeTimers();
		const insert = vi.fn();
		const tab = make({ insert });
		choose(tab, 'insert.illustrations.shapes', SHAPE_PRESET_DEFS[2].type);
		vi.useRealTimers();
		expect(insert).toHaveBeenCalledWith('shape', SHAPE_PRESET_DEFS[2].type);
	});

	it('calls insertImage() and insertMedia() for their buttons', () => {
		const insertImage = vi.fn();
		const insertMedia = vi.fn();
		const tab = make({ insertImage, insertMedia });
		pick(tab, 'insert.images.pictures', 'thisDevice');
		pick(tab, 'insert.media.media', 'videoThisDevice');
		expect(insertImage).toHaveBeenCalledOnce();
		expect(insertMedia).toHaveBeenCalledOnce();
	});

	it('inserts the chart chosen from the Chart gallery, including Bar and Pareto', () => {
		vi.useFakeTimers();
		const insertChart = vi.fn();
		const tab = make({ insertChart });
		// The gallery carries the entry id: Column and Bar share the 'bar' family.
		for (const id of [INSERT_CHART_TYPES[0].id, 'bar', 'pareto']) {
			choose(tab, 'insert.illustrations.chart', id);
			expect(insertChart).toHaveBeenLastCalledWith(id);
		}
		vi.useRealTimers();
	});

	it('opens the Header & Footer dialog and the equation panel', () => {
		const headerFooter = vi.fn();
		const equation = vi.fn();
		const tab = make({}, { headerFooter, equation });
		control(tab, 'insert.text.headerFooter').click();
		tab.el
			.querySelector<HTMLButtonElement>('[data-ribbon-control="insert.symbols.equation"] .primary')!
			.click();
		expect(headerFooter).toHaveBeenCalledOnce();
		expect(equation).toHaveBeenCalledOnce();
	});

	it('opens the shared Insert SmartArt dialog and inserts the chosen layout', () => {
		const insertSmartArt = vi.fn();
		const tab = make({ insertSmartArt });
		control(tab, 'insert.illustrations.smartArt').click();
		expect(shared(SMARTART).open).toBe(true);
		const root = pickerRoot(SMARTART, 'office-ui-smartart-picker');
		expect(root.querySelector('[role="listbox"][aria-label="SmartArt layouts"]')).not.toBeNull();
		const category = CATEGORIES[1];
		root.querySelector<HTMLButtonElement>(`[data-category="${category.id}"]`)!.click();
		const expected = PRESETS.filter((preset) => preset.category === category.id);
		expect(root.querySelectorAll('[role="option"]')).toHaveLength(expected.length);
		footerButton(SMARTART, 'insert').click();
		expect(insertSmartArt).toHaveBeenCalledWith(expected[0].layout, expected[0].defaultItems);
		expect(shared(SMARTART).open).toBe(false);
	});

	it('cancels SmartArt insertion and returns focus to the SmartArt control', async () => {
		const insertSmartArt = vi.fn();
		const tab = make({ insertSmartArt });
		control(tab, 'insert.illustrations.smartArt').click();
		footerButton(SMARTART, 'cancel').click();
		await Promise.resolve();
		expect(insertSmartArt).not.toHaveBeenCalled();
		expect(shared(SMARTART).open).toBe(false);
		expect(tab.el.querySelector('[data-ribbon-control="insert.illustrations.smartArt"]')).toBe(
			document.activeElement,
		);
	});

	it('opens the shared Insert Chart dialog from the Chart menu and inserts its choice', () => {
		const insertChart = vi.fn();
		const tab = make({ insertChart });
		pick(tab, 'insert.illustrations.chart', '__chartDialog__');
		expect(shared(CHART).open).toBe(true);
		const picker = shared(CHART).querySelector('office-ui-chart-picker') as HTMLElement & {
			value: { type: string; grouping: string; title: string };
		};
		picker.value = { type: 'bar', grouping: 'stacked', title: 'Sales' };
		footerButton(CHART, 'insert').click();
		expect(insertChart).toHaveBeenCalledExactlyOnceWith('bar', {
			title: 'Sales',
			grouping: 'stacked',
		});
		expect(shared(CHART).open).toBe(false);
	});

	it('dispatches insertActionButton(type) and insertField(type) from the shared menus', () => {
		const insertActionButton = vi.fn();
		const insertField = vi.fn();
		const tab = make({ insertActionButton, insertField });
		tab.el
			.querySelector<HTMLButtonElement>('[data-ribbon-control="insert.links.action"] .trigger')!
			.click();
		tab.el
			.querySelector<HTMLButtonElement>(
				'[data-ribbon-control="insert.links.action"] [role=menuitem]',
			)!
			.click();
		control(tab, 'insert.text.slideNumber').click();
		expect(insertActionButton).toHaveBeenCalledOnce();
		expect(insertField).toHaveBeenCalledWith('slidenum');
	});

	it('inserts a table of the size picked on the Table grid', () => {
		const insert = vi.fn();
		const tab = make({ insert });
		const table = '[data-ribbon-control="insert.tables.table"]';
		tab.el.querySelector<HTMLButtonElement>(`${table} .trigger`)!.click();
		tab.el.querySelector<HTMLElement>(`${table} [data-row="4"][data-column="3"]`)!.click();
		expect(insert).toHaveBeenCalledExactlyOnceWith('table', undefined, { rows: 4, columns: 3 });
	});

	it('adds a slide on the active layout, or the chosen one, and opens the comments pane', () => {
		const addSlide = vi.fn();
		const insertSlideFromLayout = vi.fn();
		const comments = vi.fn();
		const tab = make({ addSlide, insertSlideFromLayout }, { comments });
		tab.setLayouts(
			[
				{ path: 'l1.xml', name: 'Title' },
				{ path: 'l2.xml', name: 'Content' },
			],
			'l2.xml',
		);
		const slide = '[data-ribbon-control="insert.slides.newSlide"]';
		tab.el.querySelector<HTMLButtonElement>(`${slide} .primary`)!.click();
		expect(insertSlideFromLayout).toHaveBeenLastCalledWith('l2.xml', 'Content');
		pick(tab, 'insert.slides.newSlide', 'l1.xml');
		expect(insertSlideFromLayout).toHaveBeenLastCalledWith('l1.xml', 'Title');
		control(tab, 'insert.comments.comment').click();
		expect(comments).toHaveBeenCalledOnce();
		tab.setLayouts([]);
		tab.el.querySelector<HTMLButtonElement>(`${slide} .primary`)!.click();
		expect(addSlide).toHaveBeenCalledOnce();
	});

	it('disables New Slide and Comment when the host wires neither', () => {
		const tab = make();
		expect(control(tab, 'insert.comments.comment').disabled).toBeTruthy();
		expect(
			tab.el.querySelector<HTMLButtonElement>(
				'[data-ribbon-control="insert.slides.newSlide"] .primary',
			)!.disabled,
		).toBeTruthy();
	});

	it('setEditable disables/enables every editing control but not Link', () => {
		const tab = make();
		const editing = ['insert.text.textBox', 'insert.text.slideNumber'];
		tab.setEditable(false);
		expect(editing.every((id) => control(tab, id).disabled)).toBeTruthy();
		for (const trigger of tab.el.querySelectorAll<HTMLButtonElement>('.trigger')) {
			expect(trigger.disabled).toBeTruthy();
		}
		tab.setEditable(true);
		expect(editing.every((id) => !control(tab, id).disabled)).toBeTruthy();
	});
});
