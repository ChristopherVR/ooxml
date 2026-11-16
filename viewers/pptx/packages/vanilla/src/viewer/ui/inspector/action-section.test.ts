import type { PptxElement } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { createTranslator } from '../../i18n';
import { createActionSection } from './action-section';
import type { InspectorHandlers, InspectorState } from './types';

const shape = (extra: Record<string, unknown> = {}): PptxElement =>
	({ type: 'shape', id: 's1', x: 0, y: 0, width: 10, height: 10, ...extra }) as PptxElement;

function build(element: PptxElement = shape()) {
	const setElementAction = vi.fn();
	const section = createActionSection(
		document,
		createTranslator(),
		() => document.createElement('div'),
		{ setElementAction } as unknown as InspectorHandlers,
	);
	section.setSlideCount(3);
	section.setCustomShows([{ id: 'cs1', name: 'Sub show' }]);
	section.update({ hasSelection: true, selectedElement: element } as InspectorState);
	return { section, setElementAction };
}

const typeSelects = (el: HTMLElement) =>
	Array.from(el.querySelectorAll<HTMLSelectElement>('[data-testid="action-type"]'));

function choose(control: HTMLSelectElement | HTMLInputElement, value: string, event = 'change') {
	control.value = value;
	control.dispatchEvent(new Event(event));
}

describe('action settings section', () => {
	it('offers React parity action labels, one select per trigger', () => {
		const { section } = build();
		const [click] = typeSelects(section.el);
		expect(typeSelects(section.el)).toHaveLength(2);
		const options = Array.from(click!.querySelectorAll('option')).map((o) => o.textContent);
		expect(options).toStrictEqual([
			'None',
			'Go to URL',
			'Go to Slide',
			'First Slide',
			'Last Slide',
			'Previous Slide',
			'Next Slide',
			'End Show',
			'Last slide viewed',
			'Custom show',
			'Open file',
			'Open presentation',
			'Play media',
			'Object action',
			'Run program',
		]);
	});

	it('waits for a URL before committing a "go to URL" action', () => {
		const { section, setElementAction } = build();
		choose(typeSelects(section.el)[0]!, 'url');
		expect(setElementAction).not.toHaveBeenCalled();
		const url = section.el.querySelector<HTMLInputElement>('[data-testid="action-url-url"]')!;
		choose(url, 'https://example.com', 'input');
		expect(setElementAction).toHaveBeenCalledWith(
			'click',
			expect.objectContaining({ trigger: 'click', type: 'url', url: 'https://example.com' }),
		);
	});

	it('commits a target-free action immediately', () => {
		const { section, setElementAction } = build();
		choose(typeSelects(section.el)[1]!, 'nextSlide');
		expect(setElementAction).toHaveBeenCalledWith(
			'hover',
			expect.objectContaining({ trigger: 'hover', type: 'nextSlide' }),
		);
	});

	it('offers the deck slides as choices and writes a 0-based index', () => {
		const { section, setElementAction } = build();
		choose(typeSelects(section.el)[0]!, 'slide');
		const slide = section.el.querySelector<HTMLSelectElement>('[data-testid="action-slide"]')!;
		expect(slide.querySelectorAll('option')).toHaveLength(4);
		choose(slide, '2');
		expect(setElementAction).toHaveBeenCalledWith(
			'click',
			expect.objectContaining({ type: 'slide', slideIndex: 2 }),
		);
	});

	it('waits for a custom show, then commits its id and returnAfter', () => {
		const { section, setElementAction } = build();
		choose(typeSelects(section.el)[0]!, 'customShow');
		expect(setElementAction).not.toHaveBeenCalled();
		const show = section.el.querySelector<HTMLSelectElement>(
			'[data-testid="pptx-action-custom-show"]',
		)!;
		choose(show, 'cs1');
		expect(setElementAction).toHaveBeenLastCalledWith(
			'click',
			expect.objectContaining({ type: 'customShow', customShowId: 'cs1' }),
		);
	});

	it('drops a half-made pick when the selection moves to another element', () => {
		const { section } = build();
		choose(typeSelects(section.el)[0]!, 'url');
		expect(section.el.querySelector('[data-testid="action-url-url"]')).not.toBeNull();
		section.update({ hasSelection: true, selectedElement: shape({ id: 's2' }) } as InspectorState);
		expect(section.el.querySelector('[data-testid="action-url-url"]')).toBeNull();
	});

	it('shows the target input for runProgram and writes into url', () => {
		const { section, setElementAction } = build();
		choose(typeSelects(section.el)[0]!, 'runProgram');
		const target = section.el.querySelector<HTMLInputElement>(
			'[data-testid="action-runProgram-url"]',
		)!;
		choose(target, 'notepad.exe', 'input');
		expect(setElementAction).toHaveBeenLastCalledWith(
			'click',
			expect.objectContaining({ type: 'runProgram', url: 'notepad.exe' }),
		);
	});
});
