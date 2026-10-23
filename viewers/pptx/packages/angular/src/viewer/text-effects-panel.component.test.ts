/**
 * TextEffectsPanelComponent: the inspector's Text Effects (text shadow, glow
 * and reflection), Angular binding.
 *
 * No Angular TestBed (see `vitest.config.ts`): components are instantiated
 * directly, inputs are stubbed as signals. Pins that the generic field renderer
 * turns the shared `TEXT_EFFECT_SECTIONS` controls into text-style patches, that
 * the panel merges them into ONE element patch (one `updateElement`, so one undo
 * step). `inspector-panel.component.test.ts` pins where the inspector mounts it.
 */
import { Injector, runInInjectionContext, signal } from '@angular/core';
import type { InputSignal, OutputEmitterRef } from '@angular/core';
import { TEXT_EFFECT_SECTIONS } from 'ooxml-ui/pptx';
import type { InspectorSectionDescriptor, InspectorSectionView } from 'ooxml-ui/pptx';
import type { PptxElement, TextStyle } from 'pptx-viewer-core';
import { describe, expect, it, vi } from 'vitest';

import { InspectorFieldsComponent } from './inspector-fields.component';
import { RecentColorsService } from './recent-colors.service';
import { TextEffectsPanelComponent } from './text-effects-panel.component';

type Target = TextStyle | undefined;
type Fields = InspectorFieldsComponent<Target, Partial<TextStyle>>;

function createFields(target: Target, recentColors?: { push: (hex: string) => void }): Fields {
	const providers = recentColors ? [{ provide: RecentColorsService, useValue: recentColors }] : [];
	const fields = runInInjectionContext(
		Injector.create({ providers }),
		() => new InspectorFieldsComponent<Target, Partial<TextStyle>>(),
	);
	Object.assign(fields, {
		sections: signal(TEXT_EFFECT_SECTIONS) as unknown as InputSignal<
			readonly InspectorSectionDescriptor<Target, Partial<TextStyle>>[]
		>,
		target: signal(target) as unknown as InputSignal<Target>,
		disabled: signal(false) as unknown as InputSignal<boolean>,
	});
	return fields;
}

function captured<T>(emitter: unknown): T[] {
	const out: T[] = [];
	vi.spyOn(emitter as OutputEmitterRef<T>, 'emit').mockImplementation((value) => {
		out.push(value);
	});
	return out;
}

function inputEvent(type: string, value: string): Event {
	const input = document.createElement('input');
	input.type = type;
	input.value = value;
	return { target: input } as unknown as Event;
}

function checkboxEvent(checked: boolean): Event {
	const box = document.createElement('pptx-ui-checkbox') as HTMLElement & { checked: boolean };
	box.checked = checked;
	return { target: box } as unknown as Event;
}

function textElement(textStyle: TextStyle): PptxElement {
	return {
		type: 'text',
		id: 'text-1',
		x: 0,
		y: 0,
		width: 200,
		height: 60,
		textStyle,
	} as PptxElement;
}

function createPanel(element: PptxElement): TextEffectsPanelComponent {
	const panel = runInInjectionContext(
		Injector.create({ providers: [] }),
		() => new TextEffectsPanelComponent(),
	);
	Object.assign(panel, { element: signal(element) as unknown as InputSignal<PptxElement> });
	return panel;
}

describe('inspectorFieldsComponent (text effects)', () => {
	it('offers the text shadow, glow and reflection toggles, with fields only once on', () => {
		const views = createFields({ textGlowColor: '#ff0000' })['views']() as InspectorSectionView[];
		expect(views.map((view) => [view.id, view.labelKey, view.on])).toStrictEqual([
			['shadow', 'pptx.textEffects.shadow', false],
			['glow', 'pptx.textEffects.glow', true],
			['reflection', 'pptx.textEffects.reflection', false],
		]);
		expect(views[1]!.fields.map((field) => field.id)).toStrictEqual(['color', 'opacity', 'size']);
		expect(views[0]!.fields).toHaveLength(0);
	});

	it('turns glow on with the defaults, and clamps a typed size', () => {
		const fields = createFields({});
		const out = captured<Partial<TextStyle>>(fields.patch);
		fields['onToggle']('glow', checkboxEvent(true));
		expect(out[0]).toStrictEqual({
			textGlowColor: '#ffff00',
			textGlowRadius: 6,
			textGlowOpacity: 0.6,
		});
		fields['onField']('glow', 'size', inputEvent('number', '400'));
		expect(out[1]).toStrictEqual({ textGlowRadius: 50 });
		fields['onField']('glow', 'size', inputEvent('number', ''));
		expect(out).toHaveLength(2);
	});

	it('pushes a committed glow colour into RecentColorsService', () => {
		const pushed: string[] = [];
		const fields = createFields({ textGlowColor: '#ff0000' }, { push: (hex) => pushed.push(hex) });
		const out = captured<Partial<TextStyle>>(fields.patch);
		fields['onColor']('glow', 'color', inputEvent('color', '#00ff00'));
		expect(out).toStrictEqual([{ textGlowColor: '#00ff00' }]);
		expect(pushed).toStrictEqual(['#00ff00']);
	});
});

describe('textEffectsPanelComponent', () => {
	it('reads the element text style and merges a glow patch into it as one element patch', () => {
		const panel = createPanel(textElement({ bold: true }));
		expect(panel['sections']).toBe(TEXT_EFFECT_SECTIONS);
		expect(panel['textStyle']()).toStrictEqual({ bold: true });
		const out = captured<Partial<PptxElement>>(panel.patch);
		panel['onEffectPatch']({ textGlowColor: '#ffff00', textGlowRadius: 6, textGlowOpacity: 0.6 });
		expect(out).toStrictEqual([
			{
				textStyle: {
					bold: true,
					textGlowColor: '#ffff00',
					textGlowRadius: 6,
					textGlowOpacity: 0.6,
				},
			},
		]);
	});

	it('glows every run as well as the body, so the slide shows it', () => {
		const element = {
			...textElement({ fontSize: 20 }),
			text: 'Hello world',
			textSegments: [
				{ text: 'Hello ', style: { fontSize: 20 } },
				{ text: 'world', style: { fontSize: 20, bold: true } },
			],
		} as PptxElement;
		const panel = createPanel(element);
		const out = captured<Partial<PptxElement> & { textSegments?: Array<{ style?: TextStyle }> }>(
			panel.patch,
		);
		const glow = { textGlowColor: '#ffff00', textGlowRadius: 6, textGlowOpacity: 0.6 };
		panel['onEffectPatch'](glow);
		expect(out).toHaveLength(1);
		expect(out[0]!.textStyle).toMatchObject({ fontSize: 20, ...glow });
		expect(out[0]!.textSegments?.map((segment) => segment.style)).toStrictEqual([
			{ fontSize: 20, ...glow },
			{ fontSize: 20, bold: true, ...glow },
		]);
	});
});
