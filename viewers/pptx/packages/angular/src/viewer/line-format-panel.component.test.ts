/**
 * LineFormatPanelComponent, Angular binding: the shared `SHAPE_LINE_SECTIONS`
 * descriptors drawn by the generic `InspectorFieldsComponent`, merged into an
 * element patch.
 *
 * No Angular TestBed (see `vitest.config.ts`): components are instantiated
 * directly, inputs are stubbed as signals.
 */
import { Injector, runInInjectionContext, signal } from '@angular/core';
import type { InputSignal, OutputEmitterRef } from '@angular/core';
import { SHAPE_LINE_SECTIONS } from 'ooxml-ui/pptx';
import type { InspectorSectionDescriptor, InspectorSectionView } from 'ooxml-ui/pptx';
import type { PptxElement, ShapeStyle } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { InspectorFieldsComponent } from './inspector-fields.component';
import { LineFormatPanelComponent } from './line-format-panel.component';

type Target = ShapeStyle | undefined;

function selectEvent(value: string): Event {
	const select = document.createElement('pptx-ui-select') as HTMLElement & { value: string };
	select.value = value;
	return { target: select } as unknown as Event;
}

describe('lineFormatPanelComponent', () => {
	const element = {
		type: 'shape',
		id: 'shape-1',
		x: 0,
		y: 0,
		width: 100,
		height: 100,
		shapeStyle: { strokeWidth: 2, lineCap: 'sq' },
	} as PptxElement;

	it('resolves the line options for the shape style', () => {
		const fields = runInInjectionContext(
			Injector.create({ providers: [] }),
			() => new InspectorFieldsComponent<Target, Partial<ShapeStyle>>(),
		);
		Object.assign(fields, {
			sections: signal(SHAPE_LINE_SECTIONS) as unknown as InputSignal<
				readonly InspectorSectionDescriptor<Target, Partial<ShapeStyle>>[]
			>,
			target: signal(element.shapeStyle) as unknown as InputSignal<Target>,
			disabled: signal(false) as unknown as InputSignal<boolean>,
		});
		const [view] = fields['views']() as InspectorSectionView[];
		expect(view!.labelKey).toBeUndefined();
		expect(view!.fields.map((f) => [f.id, f.value])).toStrictEqual([
			['strokeWidth', 2],
			['strokeOpacity', 100],
			['strokeDash', 'solid'],
			['compoundLine', 'sng'],
			['lineJoin', 'round'],
			['lineCap', 'sq'],
		]);

		const out: Partial<ShapeStyle>[] = [];
		vi.spyOn(fields.patch as OutputEmitterRef<Partial<ShapeStyle>>, 'emit').mockImplementation(
			(value) => {
				out.push(value);
			},
		);
		fields['onField']('line', 'lineJoin', selectEvent('miter'));
		expect(out).toStrictEqual([{ lineJoin: 'miter' }]);
	});

	it('merges a line patch into the element shape style', () => {
		const panel = runInInjectionContext(
			Injector.create({ providers: [] }),
			() => new LineFormatPanelComponent(),
		);
		Object.assign(panel, { element: signal(element) as unknown as InputSignal<PptxElement> });
		let emitted: Partial<PptxElement> | undefined;
		vi.spyOn(panel.patch as OutputEmitterRef<Partial<PptxElement>>, 'emit').mockImplementation(
			(value) => {
				emitted = value;
			},
		);
		expect(panel['applicable']()).toBeTruthy();
		panel['onLinePatch']({ strokeDash: 'dash' });
		expect(emitted).toStrictEqual({
			shapeStyle: { strokeWidth: 2, lineCap: 'sq', strokeDash: 'dash' },
		});
	});
});
