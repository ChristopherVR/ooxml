/**
 * The inspector's fill, line and text colour pickers preview while the user
 * drags through the native picker, and a drag is ONE undo step and ONE recent
 * colour (shared `createLiveColorInput`), as in React, Vue and vanilla.
 *
 * The Angular inspector bound only `(change)` on its colour inputs, so nothing
 * moved on the slide until the picker closed. No TestBed (see
 * `vitest.config.ts`): the real panel is constructed in an injector over a
 * real `EditorStateService`, and each DOM event runs the binding its template
 * declares for it (`template-events.test-support.ts`), so a missing `(input)`
 * binding shows up as a colour that did not move.
 */
import { computed, Injector, runInInjectionContext, signal } from '@angular/core';
import type { InputSignal, OutputEmitterRef } from '@angular/core';
import { SHAPE_EFFECT_SECTIONS } from 'ooxml-ui/pptx';
import type { InspectorSectionDescriptor } from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide, ShapeStyle } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { componentSource } from './component-source.test-support';
import { EditorStateService } from './editor-state.service';
import { EffectsPanelComponent } from './effects-panel.component';
import { InspectorFieldsComponent } from './inspector-fields.component';
import { InspectorPanelComponent } from './inspector-panel.component';
import { IsMobileService } from './is-mobile';
import { RecentColorsService } from './recent-colors.service';
import {
	colorInputEvent,
	fireTemplateEvent,
	templateEventBindings,
} from './template-events.test-support';

const ACCENT1 = { scheme: 'accent1' } as const;

function shape(): PptxElement {
	return {
		type: 'shape',
		id: 'shape-1',
		x: 0,
		y: 0,
		width: 200,
		height: 100,
		shapeType: 'rect',
		text: 'Hi',
		shapeStyle: {
			fillColor: '#4472c4',
			fillColorRef: ACCENT1,
			strokeColor: '#ed7d31',
			strokeColorRef: { scheme: 'accent2' },
			strokeWidth: 2,
		},
		textStyle: { color: '#112233', colorRef: { scheme: 'tx1' } },
	} as unknown as PptxElement;
}

type Control = 'fill' | 'line' | 'text';

const CONTROLS: Record<
	Control,
	{ marker: string; color(el: PptxElement): string | undefined; ref(el: PptxElement): unknown }
> = {
	fill: {
		marker: 'id="insp-fill"',
		color: (el) => (el as { shapeStyle?: ShapeStyle }).shapeStyle?.fillColor,
		ref: (el) => (el as { shapeStyle?: ShapeStyle }).shapeStyle?.fillColorRef,
	},
	line: {
		marker: 'id="insp-stroke"',
		color: (el) => (el as { shapeStyle?: ShapeStyle }).shapeStyle?.strokeColor,
		ref: (el) => (el as { shapeStyle?: ShapeStyle }).shapeStyle?.strokeColorRef,
	},
	text: {
		marker: 'id="insp-text-color"',
		color: (el) => (el as { textStyle?: { color?: string } }).textStyle?.color,
		ref: (el) => (el as { textStyle?: { colorRef?: unknown } }).textStyle?.colorRef,
	},
};

const panelSource = componentSource(import.meta.dirname, 'inspector-panel.component.ts');

/**
 * The real panel over a real editor. With `live`, the panel's `element` input
 * follows the editor at once (change detection between every event); without
 * it the input keeps the element it was given, as when several picker events
 * arrive before the next change-detection pass.
 */
function harness(live = true) {
	const editor = new EditorStateService();
	editor.setSlides([{ id: 's1', rId: 'r1', slideNumber: 1, elements: [shape()] } as PptxSlide]);
	const pushed: string[] = [];
	const recentColors = { recent: signal<string[]>([]), push: (hex: string) => pushed.push(hex) };
	const panel = runInInjectionContext(
		Injector.create({
			providers: [
				{ provide: EditorStateService, useValue: editor },
				{ provide: RecentColorsService, useValue: recentColors },
				{ provide: IsMobileService, useValue: { isMobile: () => false } },
			],
		}),
		() => new InspectorPanelComponent(),
	);
	const current = (): PptxElement => editor.slides()[0]!.elements[0]!;
	const original = current();
	Object.assign(panel, {
		element: live ? computed(current) : signal(original),
		slideIndex: signal(0),
		canEdit: signal(true),
	});
	const fire = (control: Control, type: 'input' | 'change' | 'blur', hex: string): void => {
		const bindings = templateEventBindings(panelSource, CONTROLS[control].marker);
		fireTemplateEvent(panel, bindings, type, colorInputEvent(type, hex));
	};
	return { editor, panel, pushed, current, fire };
}

const SHADES = ['#330000', '#660000', '#990000'];

describe.each(['fill', 'line', 'text'] as const)('the %s colour picker', (control) => {
	const { color, ref } = CONTROLS[control];

	it('shows a dragged shade on the element at once, with no undo step or recent colour', () => {
		const { editor, pushed, current, fire } = harness();
		fire(control, 'input', '#330000');

		expect(color(current())).toBe('#330000');
		expect(editor.canUndo()).toBeFalsy();
		expect(editor.dirty()).toBeFalsy();
		expect(pushed).toStrictEqual([]);
	});

	it('makes a drag and its settling change ONE undo step and ONE recent colour', () => {
		const { editor, pushed, current, fire } = harness();
		const before = current();
		for (const shade of SHADES) {
			fire(control, 'input', shade);
		}
		expect(color(current())).toBe('#990000');
		fire(control, 'change', '#990000');

		expect(color(current())).toBe('#990000');
		// A native pick is a custom colour: the theme ref goes.
		expect(ref(current())).toBeUndefined();
		expect(editor.dirty()).toBeTruthy();
		expect(pushed).toStrictEqual(['#990000']);

		editor.undo();
		expect(color(current())).toBe(color(before));
		expect(ref(current())).toStrictEqual(ref(before));
		expect(editor.canUndo()).toBeFalsy();
	});

	it('leaves no undo step, and keeps the theme colour, when a drag settles where it started', () => {
		const { editor, pushed, current, fire } = harness();
		const before = current();
		fire(control, 'input', '#330000');
		fire(control, 'input', color(before)!);
		fire(control, 'change', color(before)!);

		expect(editor.canUndo()).toBeFalsy();
		// The `change` is still a pick: its colour counts as recent, though nothing was edited.
		expect(pushed).toStrictEqual([color(before)]);
		expect(color(current())).toBe(color(before));
		expect(ref(current())).toStrictEqual(ref(before));
	});

	it('settles a drag that no change event ended when the picker loses focus', () => {
		const { editor, pushed, current, fire } = harness();
		const before = current();
		fire(control, 'input', '#330000');
		fire(control, 'blur', '#330000');

		expect(pushed).toStrictEqual(['#330000']);
		editor.undo();
		expect(color(current())).toBe(color(before));
		expect(editor.canUndo()).toBeFalsy();
	});

	it('commits a pick whose input and change arrive in one turn as one undo step', () => {
		// No change detection between the two events, as when a picker commits at once.
		const { editor, pushed, current, fire } = harness(false);
		const before = current();
		fire(control, 'input', '#00ff00');
		fire(control, 'change', '#00ff00');

		expect(color(current())).toBe('#00ff00');
		expect(pushed).toStrictEqual(['#00ff00']);
		editor.undo();
		expect(color(current())).toBe(color(before));
		expect(editor.canUndo()).toBeFalsy();
	});

	it('counts a pick of the colour already there as a recent colour, with no undo step', () => {
		const { editor, pushed, current, fire } = harness(false);
		const before = current();
		fire(control, 'input', color(before)!);
		fire(control, 'change', color(before)!);

		expect(pushed).toStrictEqual([color(before)]);
		expect(editor.canUndo()).toBeFalsy();
		expect(ref(current())).toStrictEqual(ref(before));
	});

	it('keeps a drag to one undo step while change detection has not run between events', () => {
		const { editor, current, fire } = harness(false);
		const before = current();
		for (const shade of SHADES) {
			fire(control, 'input', shade);
		}
		expect(color(current())).toBe('#990000');
		fire(control, 'change', '#990000');

		editor.undo();
		expect(color(current())).toBe(color(before));
		expect(editor.canUndo()).toBeFalsy();
	});
});

describe('the generic field renderer colour control (shape effects)', () => {
	const style: ShapeStyle = { glowColor: '#ff0000', glowRadius: 8, glowOpacity: 0.5 };
	const scope = { section: { id: 'glow' }, field: { id: 'color' } };
	const fieldsSource = componentSource(import.meta.dirname, 'inspector-fields.component.ts');

	function fields() {
		const target = signal<ShapeStyle | undefined>({ ...style });
		const pushed: string[] = [];
		const committed: Partial<ShapeStyle>[] = [];
		const previews: Partial<ShapeStyle>[] = [];
		const component = runInInjectionContext(
			Injector.create({
				providers: [
					{ provide: RecentColorsService, useValue: { push: (hex: string) => pushed.push(hex) } },
				],
			}),
			() => new InspectorFieldsComponent<ShapeStyle | undefined, Partial<ShapeStyle>>(),
		);
		Object.assign(component, {
			sections: signal(SHAPE_EFFECT_SECTIONS) as unknown as InputSignal<
				readonly InspectorSectionDescriptor<ShapeStyle | undefined, Partial<ShapeStyle>>[]
			>,
			target: target as unknown as InputSignal<ShapeStyle | undefined>,
			disabled: signal(false) as unknown as InputSignal<boolean>,
		});
		const outputs = component as unknown as Record<
			string,
			OutputEmitterRef<Partial<ShapeStyle>> | undefined
		>;
		outputs['patch']!.subscribe((patch) => {
			committed.push(patch);
			target.update((current) => ({ ...current, ...patch }));
		});
		(
			outputs['preview'] as unknown as OutputEmitterRef<{ patch: Partial<ShapeStyle> }> | undefined
		)?.subscribe(({ patch }) => {
			previews.push(patch);
			target.update((current) => ({ ...current, ...patch }));
		});
		const bindings = templateEventBindings(fieldsSource, 'class="pptx-ng-fields__color"');
		const fire = (type: 'input' | 'change' | 'blur', hex: string): void =>
			fireTemplateEvent(component, bindings, type, colorInputEvent(type, hex), scope);
		return { fire, pushed, committed, previews, target };
	}

	it('previews a dragged shade without committing it or adding a recent colour', () => {
		const { fire, pushed, committed, target } = fields();
		fire('input', '#00ff00');

		expect(target()?.glowColor).toBe('#00ff00');
		expect(committed).toStrictEqual([]);
		expect(pushed).toStrictEqual([]);
	});

	it('commits a settled drag once, from the style before the drag, with one recent colour', () => {
		const { fire, pushed, committed, previews } = fields();
		fire('input', '#00ff00');
		fire('input', '#0000ff');
		fire('change', '#0000ff');

		// The previews are put back before the one commit, so its undo step starts from red.
		// Only the leaf the drag changed goes back: the radius and opacity were never touched.
		expect(previews.at(-1)).toStrictEqual({ glowColor: '#ff0000' });
		expect(committed).toStrictEqual([{ glowColor: '#0000ff', glowRadius: 8, glowOpacity: 0.5 }]);
		expect(pushed).toStrictEqual(['#0000ff']);
	});

	it('counts a pick of the colour already there as a recent colour, without a commit', () => {
		const { fire, pushed, committed } = fields();
		fire('input', '#ff0000');
		fire('change', '#ff0000');

		expect(pushed).toStrictEqual(['#ff0000']);
		expect(committed).toStrictEqual([]);
	});

	it('reaches the slide through the effects panel and the inspector without an undo step', () => {
		const { editor, panel, current } = harness();
		const effects = runInInjectionContext(
			Injector.create({ providers: [] }),
			() => new EffectsPanelComponent(),
		);
		Object.assign(effects, { element: computed(current) });
		const effectsSource = componentSource(import.meta.dirname, 'effects-panel.component.ts');
		const toEffects = templateEventBindings(effectsSource, '<pptx-inspector-fields');
		const toInspector = templateEventBindings(panelSource, '<pptx-effects-panel');
		(effects as unknown as Record<string, OutputEmitterRef<unknown> | undefined>)[
			'preview'
		]?.subscribe((patch) => fireTemplateEvent(panel, toInspector, 'preview', patch));

		fireTemplateEvent(effects, toEffects, 'preview', {
			patch: { glowColor: '#00ff00' },
			target: current().id,
		});

		expect((current() as { shapeStyle?: ShapeStyle }).shapeStyle?.glowColor).toBe('#00ff00');
		expect(editor.canUndo()).toBeFalsy();
	});
});
