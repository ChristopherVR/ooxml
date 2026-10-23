/**
 * The inspector's Size & Position fields, Angular binding: the shared
 * `ELEMENT_GEOMETRY_SECTIONS` descriptors drawn by `pptx-inspector-fields`,
 * committed through the editor's move, resize and update operations.
 *
 * No Angular TestBed (see `vitest.config.ts`): the commit handler is called
 * against a stub, and the template wiring is guarded by its source text.
 */
import { Injector, runInInjectionContext, signal } from '@angular/core';
import type { InputSignal, OutputEmitterRef } from '@angular/core';
import { ELEMENT_GEOMETRY_SECTIONS } from 'ooxml-ui/pptx';
import type {
	ElementGeometryPatch,
	ElementGeometryTarget,
	InspectorSectionDescriptor,
	InspectorSectionView,
} from 'ooxml-ui/pptx';
import { describe, expect, it, vi } from 'vitest';

import { componentSource } from './component-source.test-support';
import { InspectorFieldsComponent } from './inspector-fields.component';
import { InspectorPanelComponent } from './inspector-panel.component';

const frame = { id: 'el-1', x: 53.4, y: 10, width: 200, height: 100, rotation: 0 };

function numberEvent(value: string): Event {
	const input = document.createElement('input');
	input.type = 'number';
	input.value = value;
	return { target: input } as unknown as Event;
}

describe('inspector Size & Position fields', () => {
	it('mounts the shared geometry descriptors in the transform section', () => {
		const source = componentSource(import.meta.dirname, 'inspector-panel.component.ts');
		expect(source).toContain('[sections]="geometrySections"');
		expect(source).toContain('(patch)="onGeometryPatch($event)"');
	});

	it('shows the frame in whole px and turns a typed width into a clamped patch', () => {
		const fields = runInInjectionContext(
			Injector.create({ providers: [] }),
			() => new InspectorFieldsComponent<ElementGeometryTarget, ElementGeometryPatch>(),
		);
		Object.assign(fields, {
			sections: signal(ELEMENT_GEOMETRY_SECTIONS) as unknown as InputSignal<
				readonly InspectorSectionDescriptor<ElementGeometryTarget, ElementGeometryPatch>[]
			>,
			target: signal(frame) as unknown as InputSignal<ElementGeometryTarget>,
			disabled: signal(false) as unknown as InputSignal<boolean>,
		});
		const [view] = fields['views']() as InspectorSectionView[];
		expect(view!.fields.map((f) => f.value)).toStrictEqual([53, 10, 200, 100, 0]);

		const out: ElementGeometryPatch[] = [];
		vi.spyOn(fields.patch as OutputEmitterRef<ElementGeometryPatch>, 'emit').mockImplementation(
			(value) => {
				out.push(value);
			},
		);
		fields['onField']('geometry', 'width', numberEvent('0'));
		expect(out).toStrictEqual([{ width: 1 }]);
	});

	it('routes a move, a resize and a rotation to the editor operations', () => {
		const editor = { setPosition: vi.fn(), resize: vi.fn(), updateElement: vi.fn() };
		const host = { el: () => frame, slideIndex: () => 2, editor };
		const commit = (
			InspectorPanelComponent.prototype as unknown as {
				onGeometryPatch(this: typeof host, patch: ElementGeometryPatch): void;
			}
		).onGeometryPatch;
		commit.call(host, { x: 222 });
		commit.call(host, { height: 40 });
		commit.call(host, { rotation: 30 });
		expect(editor.setPosition).toHaveBeenCalledWith(2, 'el-1', 222, 10);
		expect(editor.resize).toHaveBeenCalledWith(2, 'el-1', 200, 40);
		expect(editor.updateElement).toHaveBeenCalledWith(2, 'el-1', { rotation: 30 });
	});
});
