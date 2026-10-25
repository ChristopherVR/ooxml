/**
 * EffectsPanelComponent and the generic InspectorFieldsComponent it draws the
 * shared `SHAPE_EFFECT_SECTIONS` with, Angular binding.
 *
 * No Angular TestBed (see `vitest.config.ts`): components are instantiated
 * directly, inputs are stubbed as signals. Pins that the renderer resolves the
 * shared descriptors, turns control events into the descriptor patches, and
 * that the panel merges them into an element patch.
 */
import { Injector, runInInjectionContext, signal } from '@angular/core';
import type { InputSignal, OutputEmitterRef } from '@angular/core';
import { SHAPE_EFFECT_SECTIONS } from 'ooxml-ui/pptx';
import type { InspectorSectionDescriptor, InspectorSectionView } from 'ooxml-ui/pptx';
import type { PptxElement, ShapeStyle } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { EffectsPanelComponent } from './effects-panel.component';
import { InspectorFieldsComponent } from './inspector-fields.component';
import { RecentColorsService } from './recent-colors.service';

type Fields = InspectorFieldsComponent<ShapeStyle | undefined, Partial<ShapeStyle>>;

function createFields(target: ShapeStyle, recentColors?: { push: (hex: string) => void }): Fields {
	const providers = recentColors ? [{ provide: RecentColorsService, useValue: recentColors }] : [];
	const fields = runInInjectionContext(
		Injector.create({ providers }),
		() => new InspectorFieldsComponent<ShapeStyle | undefined, Partial<ShapeStyle>>(),
	);
	Object.assign(fields, {
		sections: signal(SHAPE_EFFECT_SECTIONS) as unknown as InputSignal<
			readonly InspectorSectionDescriptor<ShapeStyle | undefined, Partial<ShapeStyle>>[]
		>,
		target: signal(target) as unknown as InputSignal<ShapeStyle | undefined>,
		disabled: signal(false) as unknown as InputSignal<boolean>,
	});
	return fields;
}

function captured(fields: Fields): Partial<ShapeStyle>[] {
	const out: Partial<ShapeStyle>[] = [];
	vi.spyOn(fields.patch as OutputEmitterRef<Partial<ShapeStyle>>, 'emit').mockImplementation(
		(value) => {
			out.push(value);
		},
	);
	return out;
}

function inputEvent(type: string, value: string): Event {
	const input = document.createElement('input');
	input.type = type;
	input.value = value;
	return { target: input } as unknown as Event;
}

/** The `change` a native colour picker fires once a pick settles (its drag `input`s only preview). */
function colorChange(value: string): Event {
	return { ...inputEvent('color', value), type: 'change' } as unknown as Event;
}

function checkboxEvent(checked: boolean): Event {
	const box = document.createElement('pptx-ui-checkbox') as HTMLElement & { checked: boolean };
	box.checked = checked;
	return { target: box } as unknown as Event;
}

describe('inspectorFieldsComponent (effects)', () => {
	it('resolves every effect section, with fields only for the enabled ones', () => {
		const fields = createFields({ shadowColor: '#000000' });
		const views = fields['views']() as InspectorSectionView[];
		expect(views.map((v) => v.id)).toContain('bevel');
		expect(views[0]!.fields.map((f) => f.id)).toContain('rotateWithShape');
		expect(views[1]!.fields).toHaveLength(0);
	});

	it('emits a patch turning rotateWithShape off', () => {
		const fields = createFields({ shadowColor: '#000000' });
		const out = captured(fields);
		fields['onField']('outerShadow', 'rotateWithShape', checkboxEvent(false));
		expect(out[0]?.shadowRotateWithShape).toBeFalsy();
	});

	it('enables an effect from its toggle and clamps a typed number', () => {
		const fields = createFields({});
		const out = captured(fields);
		fields['onToggle']('glow', checkboxEvent(true));
		expect(out[0]).toStrictEqual({ glowColor: '#ffff00', glowRadius: 6, glowOpacity: 0.75 });
		fields['onField']('glow', 'radius', inputEvent('number', '400'));
		expect(out[1]?.glowRadius).toBe(96);
		fields['onField']('glow', 'radius', inputEvent('number', ''));
		expect(out).toHaveLength(2);
	});

	it('pushes a committed colour into RecentColorsService', () => {
		const pushed: string[] = [];
		const fields = createFields({ shadowColor: '#000000' }, { push: (hex) => pushed.push(hex) });
		captured(fields);
		fields['onColor']('outerShadow', 'color', colorChange('#ff00ff'));
		expect(pushed).toStrictEqual(['#ff00ff']);
	});

	it('does not throw when RecentColorsService is unavailable (standalone unit test)', () => {
		const fields = createFields({});
		captured(fields);
		expect(() => fields['onColor']('glow', 'color', colorChange('#123456'))).not.toThrow();
	});
});

describe('effectsPanelComponent', () => {
	it('merges an effect patch into the element shape style', () => {
		const panel = runInInjectionContext(Injector.create({ providers: [] }), () => {
			return new EffectsPanelComponent();
		});
		const element = {
			type: 'shape',
			id: 'shape-1',
			x: 0,
			y: 0,
			width: 100,
			height: 100,
			shapeStyle: { fillColor: '#abcdef' },
		} as PptxElement;
		Object.assign(panel, { element: signal(element) as unknown as InputSignal<PptxElement> });
		let emitted: Partial<PptxElement> | undefined;
		vi.spyOn(panel.patch as OutputEmitterRef<Partial<PptxElement>>, 'emit').mockImplementation(
			(value) => {
				emitted = value;
			},
		);
		expect(panel['shapeStyle']()).toStrictEqual({ fillColor: '#abcdef' });
		panel['onEffectPatch']({ softEdgeRadius: 4 });
		expect(emitted).toStrictEqual({ shapeStyle: { fillColor: '#abcdef', softEdgeRadius: 4 } });
	});
});
