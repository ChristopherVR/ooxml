import { describe, expect, it, vi } from 'vitest';

import { createShapeEffectsControls } from './shape-effects-controls';
import type { InspectorHandlers, InspectorState } from './types';

/**
 * Mount with the identity translator so an option's text IS the i18n key it
 * resolved, which is what distinguishes "spelled from the shared arrowhead
 * catalogue" from "printed the raw `a:tailEnd` token".
 */
function mount() {
	const setShapeStyle = vi.fn();
	const controls = createShapeEffectsControls(document, (key) => key, {
		setShapeStyle,
		setShapeType: vi.fn(),
	} as unknown as InspectorHandlers);
	controls.update({ canShape: true, isConnector: true } as InspectorState);
	const labels = Array.from(controls.el.querySelectorAll('label'));
	const arrowFor = (key: string): HTMLElementTagNameMap['pptx-ui-select'] =>
		labels.find((label) => label.textContent?.startsWith(key))!.querySelector('pptx-ui-select')!;
	return { controls, setShapeStyle, arrowFor };
}

describe('shape effects arrowhead pickers', () => {
	it('keeps the six arrowhead tokens as the option values', () => {
		const { arrowFor } = mount();

		for (const key of ['pptx.connectorArrows.startArrow', 'pptx.connectorArrows.endArrow']) {
			expect(Array.from(arrowFor(key).options).map((option) => option.value)).toStrictEqual([
				'none',
				'triangle',
				'stealth',
				'diamond',
				'oval',
				'arrow',
			]);
		}
	});

	it('spells the arrowheads rather than showing `stealth`', () => {
		const { arrowFor } = mount();

		expect(
			Array.from(arrowFor('pptx.connectorArrows.startArrow').options).map(
				(option) => option.textContent,
			),
		).toStrictEqual([
			'pptx.arrowhead.none',
			'pptx.arrowhead.triangle',
			'pptx.arrowhead.stealth',
			'pptx.arrowhead.diamond',
			'pptx.arrowhead.oval',
			'pptx.arrowhead.openArrow',
		]);
	});

	it('still writes the wire token onto the connector style', () => {
		const { arrowFor, setShapeStyle } = mount();
		const start = arrowFor('pptx.connectorArrows.startArrow');

		start.value = 'stealth';
		start.dispatchEvent(new Event('change'));

		expect(setShapeStyle).toHaveBeenCalledWith({ connectorStartArrow: 'stealth' });
	});
});

describe('shape effects (shared descriptors)', () => {
	function mountWith(shapeStyle: Record<string, unknown>) {
		const setShapeStyle = vi.fn();
		const pushRecentColor = vi.fn();
		const controls = createShapeEffectsControls(document, (key) => key, {
			setShapeStyle,
			setShapeType: vi.fn(),
			pushRecentColor,
		} as unknown as InspectorHandlers);
		controls.update({ canShape: true, shapeStyle } as unknown as InspectorState);
		const byTestId = (id: string) =>
			controls.el.querySelector<HTMLInputElement>(`[data-testid="${id}"]`);
		return { controls, setShapeStyle, pushRecentColor, byTestId };
	}

	it('draws a toggle per effect section and hides the fields of the ones that are off', () => {
		const { byTestId } = mountWith({});
		for (const id of ['fx-outer-shadow-toggle', 'fx-glow-toggle', 'fx-bevel-toggle']) {
			expect(byTestId(id)).not.toBeNull();
		}
		expect(byTestId('fx-outer-shadow-blur')).toBeNull();
	});

	it('writes shadowRotateWithShape when toggled', () => {
		const { byTestId, setShapeStyle } = mountWith({ shadowColor: '#000000' });
		const checkbox = byTestId('fx-outer-shadow-rotate-with-shape')!;
		expect(checkbox).not.toBeNull();

		checkbox.checked = false;
		checkbox.dispatchEvent(new Event('change'));

		expect(setShapeStyle).toHaveBeenCalledWith(
			expect.objectContaining({ shadowRotateWithShape: false }),
		);
	});

	it('enables an effect from its toggle and rebuilds the section with its fields', () => {
		const { controls, byTestId, setShapeStyle } = mountWith({});
		const toggle = byTestId('fx-reflection-toggle')!;
		toggle.checked = true;
		toggle.dispatchEvent(new Event('change'));
		const patch = setShapeStyle.mock.calls[0]![0] as Record<string, number>;
		expect(patch['reflectionStartOpacity']).toBe(50);

		controls.update({ canShape: true, shapeStyle: patch } as unknown as InspectorState);
		const end = byTestId('fx-reflection-end')!;
		expect(end.value).toBe('0');
		end.value = '30';
		end.dispatchEvent(new Event('change'));
		expect(setShapeStyle).toHaveBeenLastCalledWith(
			expect.objectContaining({ reflectionEndOpacity: 30, reflectionStartOpacity: 50 }),
		);
	});

	it('disables every effect control when the selection cannot be shaped', () => {
		const setShapeStyle = vi.fn();
		const controls = createShapeEffectsControls(document, (key) => key, {
			setShapeStyle,
			setShapeType: vi.fn(),
		} as unknown as InspectorHandlers);
		controls.update({ canShape: false, shapeStyle: {} } as unknown as InspectorState);
		const toggle = controls.el.querySelector<HTMLInputElement>('[data-testid="fx-glow-toggle"]')!;
		expect(toggle.disabled).toBeTruthy();
	});
});
