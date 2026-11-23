import type { PptxSlide } from 'pptx-viewer-core';
import { afterEach, describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../i18n';
import { createMobileActionSheets } from './mobile-action-sheets';
import type { RibbonHandlers } from './ribbon/ribbon-types';

afterEach(() => document.body.replaceChildren());

function open(hiddenActions?: Parameters<typeof createMobileActionSheets>[5]) {
	const handlers = {
		nav: { toggleNotes: vi.fn() },
		insert: { insert: vi.fn() },
	} as unknown as RibbonHandlers;
	const source = document.createElement('pptx-ui-ribbon-insert');
	source.state = {
		editable: true,
		hasSelection: false,
		shapeType: 'rect',
		chartKind: 'column',
		newSlideAvailable: true,
	};
	const requests: unknown[] = [];
	source.addEventListener('insert-request', (event) =>
		requests.push((event as CustomEvent).detail),
	);
	const sheets = createMobileActionSheets(
		document,
		createTranslator(),
		handlers,
		vi.fn(),
		document.createElement('div'),
		hiddenActions,
		() => source,
	);
	document.body.append(sheets.el);
	sheets.update(0, [{ id: 's1' }, { id: 's2' }] as unknown as PptxSlide[], []);
	const buttons = () =>
		Array.from(
			sheets.el.querySelector('pptx-ui-mobile-bar')!.shadowRoot!.querySelectorAll('button'),
		);
	return { sheets, handlers, buttons, requests };
}

describe('mobile action sheets bottom bar adapter', () => {
	it('renders the five slots in the shared bar', () => {
		const { buttons } = open();
		expect(buttons().map((button) => button.dataset.mobileAction)).toStrictEqual([
			'slides',
			'insert',
			'inspector',
			'comments',
			'notes',
		]);
	});

	it('routes Menu > Insert to the shared sheet', () => {
		const { sheets, handlers } = open();
		sheets.toggle('menu');
		const insert = [...sheets.el.querySelectorAll('button')].find(
			(button) => button.textContent === 'Insert',
		);
		insert?.click();
		expect(sheets.el.querySelector('pptx-ui-mobile-insert')).not.toBeNull();
		expect(handlers.insert.insert).not.toHaveBeenCalled();
	});

	it('toggles the notes drawer through the handlers', () => {
		const { buttons, handlers } = open();
		buttons()[4].click();
		expect(handlers.nav.toggleNotes).toHaveBeenCalledOnce();
	});

	it('opens the shared phone Insert sheet and hands a pick to the ribbon Insert element', () => {
		const { buttons, sheets, requests } = open();
		buttons()[1].click();
		expect(buttons()[1].getAttribute('aria-pressed')).toBe('true');
		const insert = sheets.el.querySelector('pptx-ui-mobile-insert')!;
		const tile = (control: string) =>
			insert.shadowRoot!.querySelector<HTMLButtonElement>(`[data-insert-control="${control}"]`)!;
		expect(
			Array.from(insert.shadowRoot!.querySelectorAll<HTMLElement>('.tile')).map(
				(item) => item.dataset.insertControl,
			),
		).toStrictEqual([
			'insert.text.textBox',
			'insert.illustrations.shapes',
			'insert.images.pictures',
			'insert.images.screenshot',
			'insert.images.photoAlbum',
			'insert.tables.table',
			'insert.illustrations.chart',
			'insert.illustrations.smartArt',
			'insert.symbols.equation',
			'insert.symbols.symbol',
			'insert.text.wordArt',
			'insert.illustrations.icons',
			'insert.illustrations.threeDModels',
			'insert.text.object',
			'insert.media.media',
			'insert.media.screenRecording',
			'insert.slides.newSlide',
			'insert.slides.reuseSlides',
			'insert.links.zoom',
			'insert.links.link',
			'insert.text.dateTime',
			'insert.comments.comment',
			'insert.camera.cameo',
			'insert.forms.forms',
			'insert.addIns.getAddIns',
			'insert.addIns.myAddIns',
		]);
		tile('insert.text.textBox').click();
		expect(requests).toStrictEqual([{ kind: 'command', value: 'textBox' }]);
		expect(buttons()[1].getAttribute('aria-pressed')).toBe('false');
	});

	it('opens a sheet from a slot, reflects it as pressed and closes it on a second tap', () => {
		const { buttons, sheets } = open();
		buttons()[0].click();
		expect(buttons()[0].getAttribute('aria-pressed')).toBe('true');
		buttons()[0].click();
		expect(buttons()[0].getAttribute('aria-pressed')).toBe('false');
		sheets.toggle('slides');
		expect(buttons()[0].getAttribute('aria-pressed')).toBe('true');
	});

	it('disables the edit-only slots in view mode and honours hiddenActions for notes', () => {
		const { buttons, sheets } = open(['notes']);
		sheets.setEditable(false);
		const state = Object.fromEntries(
			buttons().map((button) => [button.dataset.mobileAction, button.disabled]),
		);
		expect(state).toMatchObject({ slides: false, insert: true, inspector: true, comments: true });
		expect(buttons()[4].hidden).toBeTruthy();
	});

	it('pressed notes follow the expanded drawer', () => {
		const { buttons, sheets } = open();
		sheets.setNotesExpanded(true);
		expect(buttons()[4].getAttribute('aria-pressed')).toBe('true');
	});
});
