import { DestroyRef, Injector, runInInjectionContext, signal } from '@angular/core';
import { insertSlidePaste } from 'ooxml-ui/pptx';
import { describe, expect, it } from 'vitest';

import { SlideSorterOverlayComponent } from './slide-sorter-overlay.component';

describe('sorter selection adapter', () => {
	it('pastes a range after the destination and zooms without leaving the sorter', () => {
		const injector = Injector.create({
			providers: [{ provide: DestroyRef, useValue: { onDestroy: () => () => undefined } }],
		});
		const component = runInInjectionContext(injector, () => new SlideSorterOverlayComponent());
		Object.defineProperty(component, 'slides', {
			configurable: true,
			value: signal(
				['a', 'b', 'c'].map((id, index) => ({ id, rId: id, slideNumber: index + 1, elements: [] })),
			),
		});
		Object.defineProperty(component, 'canEdit', { value: signal(true) });
		component.onThumbClick(0, new MouseEvent('click'));
		component.onThumbClick(1, new MouseEvent('click', { shiftKey: true }));
		const duplicates: number[] = [];
		component.duplicateSlide.subscribe((index) => duplicates.push(index));
		const pastes: { insertAt: number; sourceIds: string[] }[] = [];
		component.pasteSlides.subscribe((request) => {
			pastes.push(request);
			Object.defineProperty(component, 'slides', {
				value: signal(insertSlidePaste(component.slides(), request)),
				configurable: true,
			});
		});
		component.onKeydown(new KeyboardEvent('keydown', { key: 'c', ctrlKey: true }));
		component.onKeydown(new KeyboardEvent('keydown', { key: 'v', ctrlKey: true }));
		expect(duplicates).toStrictEqual([]);
		expect(pastes).toStrictEqual([expect.objectContaining({ insertAt: 2, sourceIds: ['a', 'b'] })]);
		component.onKeydown(new KeyboardEvent('keydown', { key: '+', ctrlKey: true }));
		expect(component.state().zoom).toBe(110);
		const escape = new KeyboardEvent('keydown', { key: 'Escape', cancelable: true });
		escape.preventDefault();
		component.onKeydown(escape);
		expect(component.state().selectedIds).toStrictEqual(['a']);
	});
});
