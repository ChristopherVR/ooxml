import { animationsLabel } from '../render';
import type { RibbonAnimationsViewState } from '../render';
import {
	isRibbonAnimationStart,
	RIBBON_ANIMATION_START_MODES,
	ribbonAnimationTiming,
} from '../render/ribbon-animation-editing';
import type { RibbonAnimationTimingEdit } from '../render/ribbon-animation-editing';

const START_LABELS = {
	onClick: ['pptx.animations.onClick', 'On Click'],
	withPrevious: ['pptx.animations.withPrevious', 'With Previous'],
	afterPrevious: ['pptx.animations.afterPrevious', 'After Previous'],
} as const;

let instance = 0;

export function createAnimationsTimingView(
	doc: Document,
	request: (edit: RibbonAnimationTimingEdit) => void,
) {
	const el = doc.createElement('div');
	el.className = 'timing';
	instance += 1;

	const label = doc.createElement('label');
	const start = doc.createElement('pptx-ui-select');
	start.id = `pptx-animations-start-${instance}`;
	start.dataset.ribbonControl = 'animations.timing.start';
	label.htmlFor = start.id;
	start.addEventListener('change', () => {
		if (!start.disabled && isRibbonAnimationStart(start.value)) {
			request({ field: 'start', value: start.value });
		}
	});
	el.append(label, start);

	const field = (name: 'duration' | 'delay', min: string) => {
		const input = doc.createElement('input');
		const caption = doc.createElement('label');
		let modelValue = '';
		let modelElementId: string | undefined;
		let modelSlideId: string | undefined;
		input.id = `pptx-animations-${name}-${instance}`;
		caption.htmlFor = input.id;
		input.type = 'number';
		input.min = min;
		input.max = '10';
		input.step = '0.001';
		input.dataset.ribbonControl = `animations.timing.${name}`;
		input.addEventListener('change', () => {
			if (!input.disabled) {
				request({ field: name, value: input.valueAsNumber * 1000 });
				input.value = modelValue;
			}
		});
		input.addEventListener('keydown', (event) => {
			if (event.key === 'Enter') {
				event.stopPropagation();
				input.blur();
			}
		});
		el.append(caption, input);
		return {
			caption,
			sync(
				value: string,
				enabled: boolean,
				elementId: string | undefined,
				slideId: string | undefined,
			) {
				input.disabled = !enabled;
				if (value !== modelValue || elementId !== modelElementId || slideId !== modelSlideId) {
					modelValue = value;
					modelElementId = elementId;
					modelSlideId = slideId;
					input.value = value;
				}
			},
		};
	};
	const duration = field('duration', '0.1');
	const delay = field('delay', '0');

	const reorder = doc.createElement('div');
	reorder.className = 'timing-reorder stack';
	reorder.dataset.ribbonControl = 'animations.timing.reorder';
	const move = (value: -1 | 1) => {
		const button = doc.createElement('button');
		button.type = 'button';
		button.addEventListener('click', () => request({ field: 'move', value }));
		reorder.append(button);
		return button;
	};
	const earlier = move(-1);
	const later = move(1);
	el.append(reorder);

	const sync = (state: RibbonAnimationsViewState) => {
		const timing = state.editable ? ribbonAnimationTiming(state.editing) : undefined;
		const text = (key: string, fallback: string) => animationsLabel(state, key, fallback);
		const mode = timing?.start ?? 'onClick';
		const values = isRibbonAnimationStart(mode)
			? RIBBON_ANIMATION_START_MODES
			: [mode, ...RIBBON_ANIMATION_START_MODES];
		const options = values.map((value) => {
			const option = doc.createElement('option');
			option.value = value;
			option.disabled = !isRibbonAnimationStart(value);
			const [key, fallback] = isRibbonAnimationStart(value)
				? START_LABELS[value]
				: [`pptx.animation.trigger.${value}`, value];
			option.textContent = text(key, fallback);
			return option;
		});
		if (
			start.options.length !== options.length ||
			options.some((option, index) => !option.isEqualNode(start.options[index]))
		) {
			start.replaceChildren(...options);
			start.value = mode;
		} else if (start.value !== mode) {
			start.value = mode;
		}
		start.disabled = !timing;
		label.textContent = text('pptx.animations.start', 'Start');
		start.setAttribute('aria-label', label.textContent);
		duration.caption.textContent = text('pptx.animations.duration', 'Duration');
		delay.caption.textContent = text('pptx.animations.delay', 'Delay');
		const { elementId, slide } = state.editing ?? {};
		duration.sync(
			timing ? String(timing.durationMs / 1000) : '0.5',
			Boolean(timing),
			elementId,
			slide?.id,
		);
		delay.sync(timing ? String(timing.delayMs / 1000) : '0', Boolean(timing), elementId, slide?.id);
		earlier.textContent = text('pptx.animations.moveEarlier', 'Move earlier');
		later.textContent = text('pptx.animations.moveLater', 'Move later');
		earlier.disabled = !timing?.earlier;
		later.disabled = !timing?.later;
	};
	return { el, sync };
}
