import { mount } from '@vue/test-utils';
import { registerPptxWebControls } from 'ooxml-ui/pptx';
import { beforeAll, describe, expect, it, vi } from 'vitest';

import { footerAction } from '../dialog-footer.test-support';
import InsertSection from './InsertSection.vue';

vi.mock(import('vue-i18n'), () => ({
	useI18n: () => ({ t: (key: string) => key }),
}));
beforeAll(registerPptxWebControls);

function mountInsert(overrides: Record<string, unknown> = {}) {
	const props = {
		canEdit: true,
		hasSelection: false,
		onOpenHyperlinkDialog: vi.fn(),
		newShapeType: 'rect',
		onSetNewShapeType: vi.fn(),
		onAddTextBox: vi.fn(),
		onAddShape: vi.fn(),
		onAddTable: vi.fn(),
		onAddChart: vi.fn(),
		onAddSmartArt: vi.fn(),
		onAddEquation: vi.fn(),
		onAddActionButton: vi.fn(),
		onInsertField: vi.fn(),
		onOpenHeaderFooter: vi.fn(),
		onOpenImagePicker: vi.fn(),
		onOpenMediaPicker: vi.fn(),
		...overrides,
	};
	return { props, wrapper: mount(InsertSection, { attachTo: document.body, props }) };
}
function button(root: Element, id: string): HTMLButtonElement {
	return root.querySelector(`[data-ribbon-control="${id}"]`)!.shadowRoot!.querySelector('button')!;
}

describe('insert section slide templates', () => {
	const menu = () =>
		document.body.querySelector<HTMLElement>('[data-ribbon-control="insert.slides.newSlide"]')!;

	it('ends the New Slide menu with Slide Templates, even for a deck with no layouts', async () => {
		const { wrapper } = mountInsert({ onInsertSlideFromTemplate: vi.fn(), layoutOptions: [] });
		menu().querySelector<HTMLButtonElement>('.trigger')!.click();
		const items = [...menu().querySelectorAll<HTMLButtonElement>('[data-insert-item]')];
		expect(items.map((item) => item.textContent)).toStrictEqual(['Slide Templates...']);
		items[0].click();
		await wrapper.vm.$nextTick();
		expect(document.body.querySelector('[role="dialog"]:not([hidden])')).not.toBeNull();
		wrapper.unmount();
	});

	it('leaves the entry out without a template handler', () => {
		const { wrapper } = mountInsert({ layoutOptions: [] });
		expect(menu().querySelector<HTMLButtonElement>('.trigger')!.disabled).toBeTruthy();
		wrapper.unmount();
	});
});

describe('insert section', () => {
	it('routes commands to the native handlers', () => {
		const { props, wrapper } = mountInsert({ hasSelection: true });
		for (const id of [
			'insert.text.textBox',
			'insert.illustrations.smartArt',
			'insert.links.link',
		]) {
			button(document.body, id).click();
		}
		document.body
			.querySelector<HTMLButtonElement>('[data-ribbon-control="insert.symbols.equation"] .primary')!
			.click();
		for (const [id, item] of [
			['insert.tables.table', 'insertTable'],
			['insert.images.pictures', 'thisDevice'],
			['insert.media.media', 'videoThisDevice'],
		]) {
			document.body
				.querySelector<HTMLButtonElement>(`[data-ribbon-control="${id}"] .trigger`)!
				.click();
			document.body
				.querySelector<HTMLButtonElement>(
					`[data-ribbon-control="${id}"] [data-insert-item="${item}"]`,
				)!
				.click();
		}
		expect(props.onAddTextBox).toHaveBeenCalledOnce();
		expect(props.onAddTable).toHaveBeenCalledOnce();
		expect(props.onOpenImagePicker).toHaveBeenCalledOnce();
		expect(props.onOpenMediaPicker).toHaveBeenCalledOnce();
		expect(props.onAddSmartArt).toHaveBeenCalledOnce();
		expect(props.onAddEquation).toHaveBeenCalledOnce();
		expect(props.onOpenHyperlinkDialog).toHaveBeenCalledOnce();
		wrapper.unmount();
	});

	it('routes the shape and chart pickers and the Action menu', () => {
		const { props, wrapper } = mountInsert();
		vi.useFakeTimers();
		// A gallery pick stages the type, then inserts it once the host has applied it.
		const pick = (control: string, value: string) => {
			document.body
				.querySelector<HTMLButtonElement>(`[data-ribbon-control="${control}"] .trigger`)!
				.click();
			document.body
				.querySelector<HTMLButtonElement>(
					`[data-ribbon-control="${control}"] [data-insert-item="${value}"]`,
				)!
				.click();
		};
		pick('insert.illustrations.shapes', 'star5');
		expect(props.onSetNewShapeType).toHaveBeenCalledExactlyOnceWith('star5');
		pick('insert.illustrations.chart', 'pie');
		vi.runAllTimers();
		vi.useRealTimers();
		expect(props.onAddShape).toHaveBeenCalledOnce();
		expect(props.onAddChart).toHaveBeenCalledOnce();
		document.body
			.querySelector<HTMLButtonElement>('[data-ribbon-control="insert.links.action"] .trigger')!
			.click();
		document.body
			.querySelector<HTMLButtonElement>(
				'[data-ribbon-control="insert.links.action"] [data-insert-item]',
			)!
			.click();
		expect(props.onAddActionButton).toHaveBeenCalledOnce();
		wrapper.unmount();
	});

	it('inserts simple fields directly and opens the native Date/Time dialog', async () => {
		const { props, wrapper } = mountInsert();
		button(document.body, 'insert.text.slideNumber').click();
		expect(props.onInsertField).toHaveBeenCalledExactlyOnceWith('slidenum');
		button(document.body, 'insert.text.dateTime').click();
		await wrapper.vm.$nextTick();
		expect(footerAction(document.body, 'insert')).toBeDefined();
		footerAction(document.body, 'insert')!.click();
		expect(props.onInsertField).toHaveBeenLastCalledWith('datetime', expect.any(String));
		wrapper.unmount();
	});

	it('reflects read-only and selection gating without invoking edit handlers', async () => {
		const { props, wrapper } = mountInsert({ canEdit: false, hasSelection: true });
		const textBox = button(document.body, 'insert.text.textBox');
		expect(textBox.disabled).toBeTruthy();
		textBox.click();
		expect(props.onAddTextBox).not.toHaveBeenCalled();
		expect(button(document.body, 'insert.links.link').disabled).toBeFalsy();
		await wrapper.setProps({ canEdit: true, hasSelection: false });
		expect(button(document.body, 'insert.text.textBox').disabled).toBeFalsy();
		expect(button(document.body, 'insert.links.link').disabled).toBeTruthy();
		wrapper.unmount();
	});

	it('hides controls the host does not support', () => {
		const { wrapper } = mountInsert({
			onAddChart: undefined,
			onInsertField: undefined,
			onOpenHeaderFooter: undefined,
		});
		const hidden = (selector: string) => document.body.querySelector<HTMLElement>(selector)!.hidden;
		expect(hidden('[data-ribbon-control="insert.illustrations.chart"]')).toBeTruthy();
		expect(hidden('[data-ribbon-control="insert.text.dateTime"]')).toBeTruthy();
		expect(hidden('[data-ribbon-control="insert.text.slideNumber"]')).toBeTruthy();
		wrapper.unmount();
	});

	it('inserts a table of the size picked on the Table grid', () => {
		const { props, wrapper } = mountInsert();
		const table = '[data-ribbon-control="insert.tables.table"]';
		document.body.querySelector<HTMLButtonElement>(`${table} .trigger`)!.click();
		document.body.querySelector<HTMLElement>(`${table} [data-row="4"][data-column="3"]`)!.click();
		expect(props.onAddTable).toHaveBeenCalledExactlyOnceWith(4, 3);
		wrapper.unmount();
	});

	it('adds a slide from the layout and opens the comments pane without toggling it shut', async () => {
		const onInsertSlideFromLayout = vi.fn();
		const onToggleComments = vi.fn();
		const layoutOptions = [
			{ path: 'l1.xml', name: 'Title' },
			{ path: 'l2.xml', name: 'Content' },
		];
		const { wrapper } = mountInsert({
			layoutOptions,
			currentLayoutPath: 'l2.xml',
			onInsertSlideFromLayout,
			onToggleComments,
		});
		const slide = '[data-ribbon-control="insert.slides.newSlide"]';
		document.body.querySelector<HTMLButtonElement>(`${slide} .primary`)!.click();
		expect(onInsertSlideFromLayout).toHaveBeenLastCalledWith('l2.xml', 'Content');
		document.body.querySelector<HTMLButtonElement>(`${slide} .trigger`)!.click();
		document.body.querySelector<HTMLButtonElement>(`${slide} [data-insert-item="l1.xml"]`)!.click();
		expect(onInsertSlideFromLayout).toHaveBeenLastCalledWith('l1.xml', 'Title');
		button(document.body, 'insert.comments.comment').click();
		expect(onToggleComments).toHaveBeenCalledOnce();
		await wrapper.setProps({ isCommentsPanelOpen: true });
		button(document.body, 'insert.comments.comment').click();
		expect(onToggleComments).toHaveBeenCalledOnce();
		wrapper.unmount();
	});

	it('leaves New Slide and Comment disabled when the host offers neither', () => {
		const { wrapper } = mountInsert();
		expect(button(document.body, 'insert.comments.comment').disabled).toBeTruthy();
		expect(
			document.body.querySelector<HTMLButtonElement>(
				'[data-ribbon-control="insert.slides.newSlide"] .primary',
			)!.disabled,
		).toBeTruthy();
		wrapper.unmount();
	});
});

describe('insert section phone variant', () => {
	it('draws the shared phone sheet, routes a tile to the same handler and reports done', () => {
		const onRequestHandled = vi.fn();
		const { props, wrapper } = mountInsert({ phone: true, onRequestHandled });
		const sheet = wrapper.element.parentElement!.querySelector('pptx-ui-mobile-insert')!;
		expect(document.body.querySelector('pptx-ui-ribbon-insert')).toBeNull();
		sheet
			.shadowRoot!.querySelector<HTMLButtonElement>('[data-insert-control="insert.text.textBox"]')!
			.click();
		expect(props.onAddTextBox).toHaveBeenCalledOnce();
		expect(onRequestHandled).toHaveBeenCalledOnce();
		wrapper.unmount();
	});
});
