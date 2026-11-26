// @vitest-environment jsdom
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import { PptxAnimationWriteService, PptxNativeAnimationService } from 'ooxml-core/pptx';

import {
	canRequestAnimations,
	EMPHASIS_PRESET_VALUES,
	ENTRANCE_PRESET_VALUES,
	EXIT_PRESET_VALUES,
	MOTION_PATH_PRESETS,
} from '../render';
import { registerPptxWebControls } from './index';
import {
	editRibbonAnimationTiming,
	ribbonAnimationTiming,
} from '../render/ribbon-animation-editing';

beforeAll(registerPptxWebControls);
afterEach(() => document.body.replaceChildren());
function mount(state: Partial<HTMLElementTagNameMap['pptx-ui-ribbon-animations']['state']> = {}) {
	const host = document.createElement('pptx-ui-ribbon-animations');
	host.state = { editable: true, hasSelection: true, ...state };
	document.body.append(host);
	return host;
}
function button(host: HTMLElement, id: string) {
	return host
		.querySelector(`[data-ribbon-control="${id}"]`)!
		.shadowRoot!.querySelector<HTMLButtonElement>('button')!;
}
const intents = (request: ReturnType<typeof vi.fn>) =>
	request.mock.calls.map(([event]) => event.detail);

describe('shared Animations view', () => {
	it('keeps every public group and control id exactly once', () => {
		const host = mount();
		for (const id of [
			'animations.preview',
			'animations.animation',
			'animations.motionPath',
			'animations.advancedAnimation',
			'animations.timing',
		]) {
			expect(host.querySelectorAll(`[data-ribbon-group="${id}"]`)).toHaveLength(1);
		}
		for (const id of [
			'animations.preview.preview',
			'animations.animation.gallery',
			'animations.animation.effectOptions',
			'animations.motionPath.gallery',
			'animations.advancedAnimation.addAnimation',
			'animations.advancedAnimation.animationPane',
			'animations.advancedAnimation.trigger',
			'animations.advancedAnimation.animationPainter',
			'animations.advancedAnimation.remove',
			'animations.timing.start',
			'animations.timing.duration',
			'animations.timing.delay',
			'animations.timing.reorder',
		]) {
			expect(host.querySelectorAll(`[data-ribbon-control="${id}"]`)).toHaveLength(1);
		}
		expect(host.querySelectorAll('[data-animation-preset]')).toHaveLength(
			ENTRANCE_PRESET_VALUES.length +
				EMPHASIS_PRESET_VALUES.length +
				EXIT_PRESET_VALUES.length +
				MOTION_PATH_PRESETS.length,
		);
	});

	it('dispatches one typed intent per command, preset and motion path', () => {
		const host = mount();
		const request = vi.fn();
		host.addEventListener('animations-request', request);
		button(host, 'animations.preview.preview').click();
		button(host, 'animations.advancedAnimation.addAnimation').click();
		button(host, 'animations.advancedAnimation.remove').click();
		host.querySelector<HTMLButtonElement>('[data-animation-preset="flyIn"]')!.click();
		host.querySelector<HTMLButtonElement>('[data-animation-preset="lineRight"]')!.click();
		host.querySelectorAll('pptx-ui-ribbon-command')[3].shadowRoot!.querySelector('button')!.click();
		expect(intents(request)).toStrictEqual([
			{ kind: 'command', value: 'preview' },
			{ kind: 'add', group: 'exit', preset: 'fadeOut' },
			{ kind: 'command', value: 'remove' },
			{ kind: 'add', group: 'entrance', preset: 'flyIn' },
			{ kind: 'add', group: 'motionPath', preset: 'lineRight' },
			{ kind: 'add', group: 'motionPath', preset: 'lineRight' },
		]);
	});

	it('gates selection-dependent controls and keeps the Animation Pane available', () => {
		const host = mount({ hasSelection: false, paneOpen: true });
		const request = vi.fn();
		host.addEventListener('animations-request', request);
		expect(button(host, 'animations.preview.preview').disabled).toBeTruthy();
		expect(button(host, 'animations.animation.effectOptions').disabled).toBeTruthy();
		expect(host.querySelector<HTMLButtonElement>('[data-animation-preset]')!.disabled).toBeTruthy();
		const pane = button(host, 'animations.advancedAnimation.animationPane');
		expect(pane.disabled).toBeFalsy();
		expect(pane.getAttribute('aria-pressed')).toBe('true');
		pane.click();
		expect(intents(request)).toStrictEqual([{ kind: 'command', value: 'animationPane' }]);
		host.state = { editable: true, hasSelection: true, previewActive: true };
		expect(pane.getAttribute('aria-pressed')).toBe('false');
		expect(button(host, 'animations.advancedAnimation.animationPainter').disabled).toBeTruthy();
		expect(host.querySelector<HTMLInputElement>('input')!.disabled).toBeTruthy();
		expect(host.querySelector('pptx-ui-select')!.disabled).toBeTruthy();
		expect(
			host
				.querySelector('[data-ribbon-control="animations.preview.preview"]')!
				.hasAttribute('active'),
		).toBeTruthy();
	});

	it('rejects malformed or read-only intents', () => {
		const state = { editable: true, hasSelection: true };
		expect(
			canRequestAnimations(state, { kind: 'add', group: 'entrance', preset: 'bogus' }),
		).toBeFalsy();
		expect(
			canRequestAnimations(state, { kind: 'add', group: 'exit', preset: 'flyIn' }),
		).toBeFalsy();
		expect(
			canRequestAnimations(state, { kind: 'add', group: 'motionPath', preset: 'x' }),
		).toBeFalsy();
		expect(canRequestAnimations(state, { kind: 'command', value: 'x' as never })).toBeFalsy();
		const readOnly = { ...state, editable: false };
		expect(canRequestAnimations(readOnly, { kind: 'command', value: 'remove' })).toBeFalsy();
		expect(
			canRequestAnimations(readOnly, { kind: 'command', value: 'animationPane' }),
		).toBeTruthy();
		const host = mount({ editable: false });
		const request = vi.fn();
		host.addEventListener('animations-request', request);
		host.querySelector<HTMLButtonElement>('[data-animation-preset]')!.click();
		expect(request).not.toHaveBeenCalled();
	});

	it('preserves focus, instance isolation and reconnect without duplicates', () => {
		const first = mount(),
			second = mount({ hasSelection: false });
		const preset = first.querySelector<HTMLButtonElement>('[data-animation-preset="fadeIn"]')!;
		preset.focus();
		first.state = { ...first.state, paneOpen: true };
		expect(document.activeElement).toBe(preset);
		expect(
			second.querySelector<HTMLButtonElement>('[data-animation-preset="fadeIn"]')!.disabled,
		).toBeTruthy();
		first.remove();
		document.body.append(first);
		expect(first.querySelectorAll('[data-ribbon-group="animations.preview"]')).toHaveLength(1);
		const request = vi.fn();
		first.addEventListener('animations-request', request);
		preset.click();
		expect(request).toHaveBeenCalledOnce();
	});

	it('uses translations when supplied and readable fallbacks otherwise', () => {
		const host = mount({
			translate: (key) => (key === 'pptx.animations.preview' ? 'Vorschau' : key),
		});
		expect(
			host.querySelector('[data-ribbon-group="animations.preview"]')!.getAttribute('label'),
		).toBe('Vorschau');
		expect(host.querySelector('[data-animation-preset="fadeIn"]')!.textContent).toBe('Fade In');
	});
});

describe('shared Animations timing', () => {
	it('selects the exact effect from its click group and labels it with the shape name', () => {
		const rawTiming = new PptxAnimationWriteService().buildTimingXml(
			[
				{ elementId: '2', entrance: 'fadeIn', durationMs: 500 },
				{ elementId: '2', exit: 'fadeOut', durationMs: 1000 },
			],
			undefined,
		)!;
		delete rawTiming['p:extLst'];
		const nativeAnimations = new PptxNativeAnimationService().parseNativeAnimations({
			'p:sld': { 'p:timing': rawTiming },
		})!;
		const host = mount({
			paneOpen: true,
			editing: {
				slide: {
					rawTiming,
					nativeAnimations,
					elements: [
						{
							id: '2',
							type: 'shape',
							name: 'Logo',
							x: 0,
							y: 0,
							width: 20,
							height: 20,
							shapeType: 'rect',
							shapeStyle: {},
						},
					],
					animationTimelineAnchors: [{ order: 1, targetIds: ['2'], presetClasses: ['exit'] }],
				},
				elementId: '2',
				commit: vi.fn(),
			},
		});
		const viewer = document.createElement('div');
		viewer.setAttribute('data-pptx-theme-root', '');
		const rows = document.createElement('pptx-ui-animation-effect-rows');
		rows.setAttribute('group-order', '2');
		viewer.append(host, rows);
		document.body.append(viewer);

		expect(rows.shadowRoot?.querySelector('button')).toBeNull();
		Object.assign(rows, { groupOrder: 1 });
		expect(rows.getAttribute('group-order')).toBe('1');
		const effect = rows.shadowRoot!.querySelector('button')!;
		expect(effect.textContent).toContain('Logo · Fade');
		effect.click();
		expect(effect.getAttribute('aria-pressed')).toBe('true');
		expect(
			host.querySelector<HTMLInputElement>('[data-ribbon-control="animations.timing.duration"]')!
				.value,
		).toBe('1');
		host.state = {
			...host.state,
			editing: {
				...host.state.editing!,
				slide: { ...host.state.editing!.slide, nativeAnimations: [...nativeAnimations] },
			},
		};
		expect(effect.getAttribute('aria-pressed')).toBe('false');
	});

	it('does not edit an authored animation for a stale native selection', () => {
		const commit = vi.fn();
		const editing = {
			slide: { animations: [{ elementId: '2', entrance: 'fadeIn' as const }] },
			elementId: '2',
			nativeEffectId: 27,
			commit,
		};
		expect(ribbonAnimationTiming(editing)).toBeUndefined();
		expect(editRibbonAnimationTiming(editing, { field: 'duration', value: 800 })).toBe('stale');
		expect(commit).not.toHaveBeenCalled();
	});
	it('associates accessible labels with the timing controls', () => {
		const host = mount();
		const duration = host.querySelector<HTMLInputElement>(
			'[data-ribbon-control="animations.timing.duration"]',
		)!;
		expect(duration.labels?.[0]?.textContent).toBe('Duration');
		const start = host.querySelector<HTMLSelectElement>(
			'[data-ribbon-control="animations.timing.start"]',
		)!;
		expect(host.querySelector(`label[for="${start.id}"]`)!.textContent).toBe('Start');
		expect(mount().querySelector('pptx-ui-select')!.id).not.toBe(start.id);
	});

	it('reads the current timing, commits seconds as milliseconds and gates unavailable moves', () => {
		const commit = vi.fn();
		const host = mount({
			editing: {
				slide: {
					animations: [{ elementId: 'source', entrance: 'fadeIn', durationMs: 600, order: 0 }],
				},
				elementId: 'source',
				commit,
			},
		});
		const duration = host.querySelector<HTMLInputElement>(
			'[data-ribbon-control="animations.timing.duration"]',
		)!;
		const start = host.querySelector('pptx-ui-select')!;
		expect(duration.value).toBe('0.6');
		expect(start.options.map((option) => option.value)).toStrictEqual([
			'onClick',
			'withPrevious',
			'afterPrevious',
		]);
		expect(
			[...host.querySelectorAll<HTMLButtonElement>('.timing-reorder button')].every(
				(button) => button.disabled,
			),
		).toBeTruthy();
		duration.value = '1.125';
		host.state = { ...host.state, paneOpen: true };
		expect(duration.value).toBe('1.125');
		duration.dispatchEvent(new Event('change'));
		expect(commit).toHaveBeenCalledExactlyOnceWith([expect.objectContaining({ durationMs: 1125 })]);
	});

	it('only applies a painter click to canvas objects in its own viewer', () => {
		const host = mount();
		const viewer = document.createElement('div');
		viewer.setAttribute('data-pptx-theme-root', '');
		const canvas = document.createElement('div');
		canvas.setAttribute('data-pptx-viewport', '');
		const target = document.createElement('button');
		target.dataset.elementId = 'target';
		canvas.append(target);
		viewer.append(host, canvas);
		document.body.append(viewer);
		const commit = vi.fn();
		host.state = {
			...host.state,
			editing: {
				slide: { animations: [{ elementId: 'source', entrance: 'fadeIn' }] },
				elementId: 'source',
				commit,
			},
		};
		const painter = button(host, 'animations.advancedAnimation.animationPainter');
		painter.click();
		expect(painter.getAttribute('aria-pressed')).toBe('true');
		canvas.click();
		expect(commit).not.toHaveBeenCalled();
		const pointer = (type: string, x: number) =>
			new PointerEvent(type, { bubbles: true, pointerId: 1, isPrimary: true, clientX: x });
		target.dispatchEvent(pointer('pointerdown', 10));
		target.dispatchEvent(pointer('pointermove', 15));
		target.dispatchEvent(pointer('pointerup', 10));
		expect(commit).not.toHaveBeenCalled();

		target.dispatchEvent(pointer('pointerdown', 10));
		const redrawn = target.cloneNode();
		target.replaceWith(redrawn);
		redrawn.dispatchEvent(pointer('pointerup', 10));
		expect(commit).toHaveBeenCalledExactlyOnceWith([
			expect.objectContaining({ elementId: 'source' }),
			expect.objectContaining({ elementId: 'target', entrance: 'fadeIn' }),
		]);
		expect(painter.getAttribute('aria-pressed')).toBe('false');
	});
});

describe('shared Animations touch contract', () => {
	it('opts every effect button out of native generic button size resets', () => {
		const host = mount();
		const buttons = [...host.querySelectorAll<HTMLButtonElement>('[data-animation-preset]')];
		expect(buttons.length).toBeGreaterThan(50);
		expect(buttons.every((item) => item.hasAttribute('data-pptx-compact'))).toBeTruthy();
	});
});

describe('shared Animations gallery structure', () => {
	it('groups motion paths under the five families and effects under the three buckets (fallback captions)', () => {
		const host = mount({ translate: (key) => key });
		const captions = (id: string) =>
			[...host.querySelectorAll(`[data-ribbon-control="${id}"] .column`)].map((node) =>
				node.getAttribute('aria-label'),
			);
		expect(captions('animations.motionPath.gallery')).toStrictEqual([
			'Lines',
			'Arcs',
			'Turns',
			'Shapes',
			'Loops',
		]);
		expect(captions('animations.animation.gallery')).toStrictEqual([
			'Entrance',
			'Emphasis',
			'Exit',
		]);
		const gallery = host.querySelector('[data-ribbon-control="animations.motionPath.gallery"]')!;
		expect(gallery.getAttribute('role')).toBe('group');
		expect(gallery.getAttribute('aria-label')).toBe(
			'Motion Paths: Lines, Arcs, Turns, Shapes, and Loops',
		);
	});
});
