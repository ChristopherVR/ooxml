/**
 * A live colour drag in the Angular inspector stays on the element it started
 * on and never overwrites another edit (shared `createLiveColorInput`):
 *
 * - the selection moves to another shape before the pick settles: the first
 *   shape gets its colour back, and the pick does not land on the second;
 * - another hand edits a field next to the colour mid-drag (a collaborator's
 *   line width): the commit and its undo step keep that edit;
 * - the effects colour fields (one controller per field) tell a new target
 *   from a re-render of their own drag.
 *
 * No TestBed (see `vitest.config.ts`): the real panel over a real
 * `EditorStateService`, events run through the template's own bindings.
 */
import { computed, Injector, runInInjectionContext, signal } from '@angular/core';
import type { InputSignal, OutputEmitterRef } from '@angular/core';
import { SHAPE_EFFECT_SECTIONS } from 'ooxml-ui/pptx';
import type { InspectorSectionDescriptor } from 'ooxml-ui/pptx';
import type { PptxElement, PptxSlide, ShapeStyle } from 'pptx-viewer-core';
import { describe, expect, it } from 'vitest';

import { componentSource } from './component-source.test-support';
import { EditorStateService } from './editor-state.service';
import { InspectorFieldsComponent } from './inspector-fields.component';
import { InspectorPanelComponent } from './inspector-panel.component';
import { IsMobileService } from './is-mobile';
import { RecentColorsService } from './recent-colors.service';
import {
	colorInputEvent,
	fireTemplateEvent,
	templateEventBindings,
} from './template-events.test-support';

function shape(id: string, shapeStyle: ShapeStyle): PptxElement {
	return {
		type: 'shape',
		id,
		x: 0,
		y: 0,
		width: 200,
		height: 100,
		shapeType: 'rect',
		shapeStyle,
	} as unknown as PptxElement;
}

const A_STYLE: ShapeStyle = {
	fillColor: '#4472c4',
	fillColorRef: { scheme: 'accent1' },
	strokeColor: '#ed7d31',
	strokeWidth: 2,
};
const B_STYLE: ShapeStyle = {
	fillColor: '#00aa55',
	fillColorRef: { scheme: 'accent6' },
	strokeWidth: 7,
};

const panelSource = componentSource(import.meta.dirname, 'inspector-panel.component.ts');
const fillBindings = templateEventBindings(panelSource, 'id="insp-fill"');

/** The panel over an editor holding shapes `a` and `b`, showing whichever is selected. */
function harness() {
	const editor = new EditorStateService();
	editor.setSlides([
		{
			id: 's1',
			rId: 'r1',
			slideNumber: 1,
			elements: [shape('a', A_STYLE), shape('b', B_STYLE)],
		} as PptxSlide,
	]);
	const selected = signal('a');
	const pushed: string[] = [];
	const panel = runInInjectionContext(
		Injector.create({
			providers: [
				{ provide: EditorStateService, useValue: editor },
				{
					provide: RecentColorsService,
					useValue: { recent: signal([]), push: (hex: string) => pushed.push(hex) },
				},
				{ provide: IsMobileService, useValue: { isMobile: () => false } },
			],
		}),
		() => new InspectorPanelComponent(),
	);
	const element = (id: string): PptxElement =>
		editor.slides()[0]!.elements.find((candidate) => candidate.id === id)!;
	Object.assign(panel, {
		element: computed(() => element(selected())),
		slideIndex: signal(0),
		canEdit: signal(true),
	});
	const style = (id: string): ShapeStyle => (element(id) as { shapeStyle: ShapeStyle }).shapeStyle;
	const fire = (type: 'input' | 'change' | 'blur', hex: string): void =>
		fireTemplateEvent(panel, fillBindings, type, colorInputEvent(type, hex));
	return { editor, selected, pushed, style, fire };
}

describe('a fill drag the selection leaves', () => {
	it('puts the first shape back and does not land the pick on the second', () => {
		const { editor, selected, pushed, style, fire } = harness();
		fire('input', '#330000');
		selected.set('b');
		fire('change', '#330000');

		expect(style('a')).toStrictEqual(A_STYLE);
		expect(style('b')).toStrictEqual(B_STYLE);
		expect(editor.canUndo()).toBeFalsy();
		expect(pushed).toStrictEqual([]);
	});
});

describe('a fill drag the selection leaves after another edit to the shape', () => {
	it('puts the colour back on the shape as it is now, keeping that edit', () => {
		const { editor, selected, style, fire } = harness();
		fire('input', '#330000');
		editor.slides.update((slides) =>
			slides.map((slide) => ({
				...slide,
				elements: slide.elements.map((el) =>
					el.id === 'a'
						? ({ ...el, shapeStyle: { ...style('a'), strokeWidth: 5 } } as PptxElement)
						: el,
				),
			})),
		);
		selected.set('b');
		fire('change', '#330000');

		expect(style('a')).toStrictEqual({ ...A_STYLE, strokeWidth: 5 });
	});
});

describe('another edit to the dragged shape', () => {
	it('survives the commit, and the one undo step keeps it', () => {
		const { editor, style, fire } = harness();
		fire('input', '#330000');
		// A collaborator widens the line while the picker is open.
		editor.slides.update((slides) =>
			slides.map((slide) => ({
				...slide,
				elements: slide.elements.map((el) =>
					el.id === 'a'
						? ({ ...el, shapeStyle: { ...style('a'), strokeWidth: 5 } } as PptxElement)
						: el,
				),
			})),
		);
		fire('change', '#330000');

		expect(style('a').strokeWidth).toBe(5);
		expect(style('a').fillColor).toBe('#330000');
		editor.undo();
		expect(style('a')).toStrictEqual({ ...A_STYLE, strokeWidth: 5 });
	});
});

describe('an effects colour field whose target changes mid-drag', () => {
	const scope = { section: { id: 'glow' }, field: { id: 'color' } };
	const fieldsSource = componentSource(import.meta.dirname, 'inspector-fields.component.ts');

	/** The generic renderer over two shapes' styles, showing `a` until `select` moves it. */
	function fieldsHarness() {
		const styles: Record<string, ShapeStyle> = {
			a: { glowColor: '#ff0000', glowRadius: 8, glowOpacity: 0.5 },
			b: { glowColor: '#ff0000', glowRadius: 3, glowOpacity: 0.9 },
		};
		let current = 'a';
		const target = signal<ShapeStyle | undefined>(styles['a']);
		const targetKey = signal('a');
		const fields = runInInjectionContext(
			Injector.create({
				providers: [{ provide: RecentColorsService, useValue: { push: () => {} } }],
			}),
			() => new InspectorFieldsComponent<ShapeStyle | undefined, Partial<ShapeStyle>>(),
		);
		Object.assign(fields, {
			sections: signal(SHAPE_EFFECT_SECTIONS) as unknown as InputSignal<
				readonly InspectorSectionDescriptor<ShapeStyle | undefined, Partial<ShapeStyle>>[]
			>,
			target: target as unknown as InputSignal<ShapeStyle | undefined>,
			targetKey: targetKey as unknown as InputSignal<string>,
			lookup: signal((key: string) => styles[key]) as unknown as InputSignal<
				(key: string) => ShapeStyle | undefined
			>,
			disabled: signal(false) as unknown as InputSignal<boolean>,
		});
		// The parent's job: a patch goes to the element it names (the selection when none).
		const apply = (patch: Partial<ShapeStyle>, id = current): void => {
			styles[id] = { ...styles[id], ...patch };
			if (id === current) {
				target.set(styles[id]);
			}
		};
		const outputs = fields as unknown as Record<string, OutputEmitterRef<unknown>>;
		outputs['patch']!.subscribe((patch) => apply(patch as Partial<ShapeStyle>));
		outputs['preview']!.subscribe((event) => {
			const named = event as { patch: Partial<ShapeStyle>; target: string };
			apply(named.patch, named.target);
		});
		const bindings = templateEventBindings(fieldsSource, 'class="pptx-ng-fields__color"');
		return {
			styles,
			fire: (type: 'input' | 'blur', hex: string): void =>
				fireTemplateEvent(fields, bindings, type, colorInputEvent(type, hex), scope),
			select(id: string) {
				current = id;
				targetKey.set(id);
				target.set(styles[id]);
			},
		};
	}

	it('does not write the first target onto the second', () => {
		const { styles, fire, select } = fieldsHarness();
		fire('input', '#00ff00');
		fire('input', '#ff0000');
		select('b');
		fire('blur', '#ff0000');

		expect(styles['b']).toStrictEqual({ glowColor: '#ff0000', glowRadius: 3, glowOpacity: 0.9 });
		expect(styles['a']).toMatchObject({ glowColor: '#ff0000', glowRadius: 8, glowOpacity: 0.5 });
	});

	it('puts the first target back as it is now, keeping an edit made there since', () => {
		const { styles, fire, select } = fieldsHarness();
		fire('input', '#00ff00');
		styles['a'] = { ...styles['a'], glowRadius: 12 };
		select('b');
		fire('blur', '#00ff00');

		expect(styles['a']).toStrictEqual({ glowColor: '#ff0000', glowRadius: 12, glowOpacity: 0.5 });
		expect(styles['b']).toStrictEqual({ glowColor: '#ff0000', glowRadius: 3, glowOpacity: 0.9 });
	});
});
