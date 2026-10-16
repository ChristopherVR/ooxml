// @vitest-environment jsdom
import { afterEach, beforeAll, describe, expect, it } from 'vitest';

import { editorViewportFitFor, EDITOR_VIEWPORT_FIT } from '../render';
import { registerPptxWebControls } from './index';
import type { PptxUiMobileInsertElement } from './mobile-insert';

beforeAll(registerPptxWebControls);

afterEach(() => document.body.replaceChildren());

function mount() {
	const element: PptxUiMobileInsertElement = document.createElement('pptx-ui-mobile-insert');
	element.state = {
		editable: true,
		hasSelection: false,
		shapeType: 'rect',
		chartKind: 'column',
		chartAvailable: true,
		newSlideAvailable: true,
		commentAvailable: true,
		fieldAvailable: true,
		slideTemplatesAvailable: true,
		layouts: [{ path: 'ppt/slideLayouts/slideLayout2.xml', name: 'Title and Content' }],
	};
	const requests: unknown[] = [];
	element.addEventListener('insert-request', (event) =>
		requests.push((event as CustomEvent).detail),
	);
	document.body.append(element);
	const tile = (control: string) =>
		element.shadowRoot!.querySelector<HTMLButtonElement>(`[data-insert-control="${control}"]`)!;
	return { element, requests, tile };
}

describe('pptx-ui-mobile-insert', () => {
	it('emits the ribbon intents for direct tiles', () => {
		const { requests, tile } = mount();
		tile('insert.text.textBox').click();
		tile('insert.images.pictures').click();
		tile('insert.comments.comment').click();
		expect(requests).toStrictEqual([
			{ kind: 'command', value: 'textBox' },
			{ kind: 'command', value: 'image' },
			{ kind: 'command', value: 'comment' },
		]);
	});

	it('has Link and Date & Time tiles that emit the ribbon intents', () => {
		const { element, requests, tile } = mount();
		element.state = { ...element.state, hasSelection: true };
		tile('insert.links.link').click();
		tile('insert.text.dateTime').click();
		expect(requests).toStrictEqual([
			{ kind: 'command', value: 'link' },
			{ kind: 'field', value: 'datetime' },
		]);
	});

	it('ends the New Slide layouts with Slide Templates when the host offers it', () => {
		const { element, requests, tile } = mount();
		tile('insert.slides.newSlide').click();
		element
			.shadowRoot!.querySelector<HTMLButtonElement>(
				'[data-insert-control="insert.slides.slideTemplates"]',
			)!
			.click();
		expect(requests).toStrictEqual([{ kind: 'slideTemplates' }]);
	});

	it('opens the Table grid and inserts the picked size, then goes back', () => {
		const { element, requests, tile } = mount();
		tile('insert.tables.table').click();
		element.shadowRoot!.querySelector<HTMLElement>('[data-row="3"][data-column="4"]')!.click();
		expect(requests).toStrictEqual([{ kind: 'table', rows: 3, columns: 4 }]);
		element.shadowRoot!.querySelector<HTMLButtonElement>('.back')!.click();
		expect(tile('insert.text.textBox')).not.toBeNull();
	});

	it('lists shapes, chart types and layouts as sub-views', () => {
		const { element, requests, tile } = mount();
		tile('insert.illustrations.shapes').click();
		element.shadowRoot!.querySelector<HTMLButtonElement>('.shape')!.click();
		element.reset();
		tile('insert.illustrations.chart').click();
		element.shadowRoot!.querySelector<HTMLButtonElement>('.row')!.click();
		element.reset();
		tile('insert.slides.newSlide').click();
		element.shadowRoot!.querySelector<HTMLButtonElement>('.row')!.click();
		expect(requests.map((item) => (item as { kind: string }).kind)).toStrictEqual([
			'shape',
			'chart',
			'newSlide',
		]);
	});

	it('does not emit while read-only', () => {
		const { element, requests, tile } = mount();
		element.state = { ...element.state, editable: false };
		tile('insert.text.textBox').click();
		expect(requests).toStrictEqual([]);
	});
});

describe('editorViewportFitFor', () => {
	it('insets the slide 12px on a phone and keeps the desktop policy elsewhere', () => {
		expect(editorViewportFitFor(412).fitPadding).toStrictEqual({ horizontal: 12, vertical: 12 });
		expect(editorViewportFitFor(1280)).toBe(EDITOR_VIEWPORT_FIT);
		expect(editorViewportFitFor(0)).toBe(EDITOR_VIEWPORT_FIT);
	});
});
