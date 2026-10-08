import { afterEach, expect, it, vi } from 'vitest';
import { createTextDraft } from './viewer-text-draft';
import { visioDrawBounds } from 'ooxml-core/visio/ui';

afterEach(() => {
	vi.restoreAllMocks();
	vi.unstubAllGlobals();
	document.body.replaceChildren();
});
it('keeps the draft and its actions inside the visible scrolled viewport near page edges', () => {
	const viewport = document.createElement('div');
	document.body.append(viewport);
	Object.defineProperties(viewport, { clientWidth: { value: 320 }, clientHeight: { value: 240 } });
	viewport.scrollLeft = 100;
	viewport.scrollTop = 80;
	vi.spyOn(HTMLElement.prototype, 'offsetWidth', 'get').mockReturnValue(300);
	vi.spyOn(HTMLElement.prototype, 'offsetHeight', 'get').mockReturnValue(210);
	vi.stubGlobal(
		'DOMPoint',
		class {
			constructor(
				public x: number,
				public y: number,
			) {}
			matrixTransform() {
				return { x: this.x, y: this.y };
			}
		},
	);
	const svg = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
	Object.defineProperty(svg, 'getScreenCTM', { value: () => ({}) });
	viewport.append(svg);
	const draft = createTextDraft(
		viewport,
		svg,
		visioDrawBounds({ x: 310, y: 235 }, { x: 315, y: 238 }),
		vi.fn(),
		vi.fn(),
	);
	expect(draft.element.style.left).toBe('112px');
	expect(draft.element.style.top).toBe('102px');
	expect(draft.element.style.maxHeight).toBe('224px');
	draft.setBusy(true);
	expect(draft.input.disabled).toBe(true);
	expect(draft.element.querySelector('[command="apply-text-box"]')).toHaveProperty(
		'disabled',
		true,
	);
	expect(draft.element.querySelector('[command="cancel-text-box"]')).not.toHaveProperty(
		'disabled',
		true,
	);
	draft.dispose();
	expect(viewport.querySelector('[data-text-draft]')).toBeNull();
});
