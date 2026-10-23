import { describe, expect, it, vi } from 'vitest';

import {
	recentSwatches,
	themeGrids,
	themeSwatch,
	themeSwatches,
} from '../theme-color-swatch-grid.test-support';
import { createFillSection } from './fill-section';
import type { InspectorHandlers, InspectorState } from './types';

const OFFICE_THEME: Record<string, string> = {
	dk1: '#000000',
	lt1: '#ffffff',
	dk2: '#44546a',
	lt2: '#e7e6e6',
	accent1: '#4472c4',
	accent2: '#ed7d31',
	accent3: '#a5a5a5',
	accent4: '#ffc000',
	accent5: '#5b9bd5',
	accent6: '#70ad47',
	bg1: '#ffffff',
	tx1: '#000000',
	bg2: '#e7e6e6',
	tx2: '#44546a',
};

function baseState(overrides: Partial<InspectorState> = {}): InspectorState {
	return {
		hasSelection: true,
		canShape: true,
		fillColor: '#4f86ff',
		strokeColor: '#1e3a8a',
		strokeWidth: 1,
		fillOpacity: 1,
		strokeOpacity: 1,
		gradientEnabled: false,
		gradient: { type: 'linear', angle: 90, stops: [] },
		shapeStyle: undefined,
		themeColorMap: undefined,
		...overrides,
	} as unknown as InspectorState;
}

function mount() {
	const handlers = {
		setShapeFill: vi.fn(),
		setShapeStroke: vi.fn(),
		setShapeStrokeWidth: vi.fn(),
		setFillOpacity: vi.fn(),
		setStrokeOpacity: vi.fn(),
		setGradientFill: vi.fn(),
		updateGradientStop: vi.fn(),
		removeGradientStop: vi.fn(),
		addGradientStop: vi.fn(),
		setShapeStyle: vi.fn(),
		setShapeType: vi.fn(),
		pushRecentColor: vi.fn(),
	} as unknown as InspectorHandlers;
	const section = createFillSection(
		document,
		(key) => key,
		(label) => {
			const el = document.createElement('section');
			el.setAttribute('aria-label', label);
			return el;
		},
		handlers,
	);
	return { section, handlers };
}

describe('createFillSection line options', () => {
	it('draws the shared Line options and writes them as shape-style patches', () => {
		const { section, handlers } = mount();
		section.update(baseState({ shapeStyle: { strokeWidth: 2.5, lineCap: 'rnd' } }));
		const byTestId = (id: string) =>
			section.el.querySelector<HTMLInputElement>(`[data-testid="${id}"]`)!;
		expect(byTestId('line-width').value).toBe('2.5');
		expect(byTestId('line-cap').value).toBe('rnd');

		const dash = byTestId('line-dash');
		dash.value = 'sysDash';
		dash.dispatchEvent(new Event('change'));
		expect(handlers.setShapeStyle).toHaveBeenCalledWith({ strokeDash: 'sysDash' });

		const width = byTestId('line-width');
		width.value = '4';
		width.dispatchEvent(new Event('change'));
		expect(handlers.setShapeStyle).toHaveBeenLastCalledWith({ strokeWidth: 4 });
	});

	it('disables the line options when the shape cannot be formatted', () => {
		const { section } = mount();
		section.update(baseState({ canShape: false, shapeStyle: {} }));
		const width = section.el.querySelector<HTMLInputElement>('[data-testid="line-width"]')!;
		expect(width.disabled).toBeTruthy();
	});
});

describe('createFillSection pattern-fill panel', () => {
	it('is hidden until the shape is in pattern fill mode', () => {
		const { section } = mount();
		section.update(baseState());
		const panel = section.el.querySelector('.pptxv-inspector-pattern') as HTMLElement;
		expect(panel.hidden).toBeTruthy();
	});

	it('shows the preset + background controls when shapeStyle.fillMode is "pattern"', () => {
		const { section } = mount();
		section.update(
			baseState({
				shapeStyle: {
					fillPatternPreset: 'cross',
					fillPatternBackgroundColor: '#123456',
					fillMode: 'pattern',
				},
			} as unknown as Partial<InspectorState>),
		);
		const panel = section.el.querySelector('.pptxv-inspector-pattern') as HTMLElement;
		expect(panel.hidden).toBeFalsy();
		const select = panel.querySelector('pptx-ui-select')!;
		expect(select.value).toBe('cross');
	});

	it('turning the pattern toggle on writes fillMode: pattern with a default preset', () => {
		const { section, handlers } = mount();
		section.update(baseState());
		const toggle = section.el.querySelector('.pptxv-inspector-pattern')
			?.previousElementSibling as HTMLElement;
		const checkbox = toggle.querySelector('pptx-ui-checkbox') as HTMLInputElement;

		checkbox.checked = true;
		checkbox.dispatchEvent(new Event('change'));

		expect(handlers.setShapeStyle).toHaveBeenCalledWith(
			expect.objectContaining({ fillMode: 'pattern', fillPatternPreset: 'pct20' }),
		);
	});

	it('turning the pattern toggle off reverts to a solid fill', () => {
		const { section, handlers } = mount();
		section.update(
			baseState({
				shapeStyle: { fillMode: 'pattern', fillPatternPreset: 'cross' },
			} as unknown as Partial<InspectorState>),
		);
		const toggle = section.el.querySelector('.pptxv-inspector-pattern')
			?.previousElementSibling as HTMLElement;
		const checkbox = toggle.querySelector('pptx-ui-checkbox') as HTMLInputElement;

		checkbox.checked = false;
		checkbox.dispatchEvent(new Event('change'));

		expect(handlers.setShapeFill).toHaveBeenCalledWith('#4f86ff');
	});

	it('changing the preset select writes fillPatternPreset', () => {
		const { section, handlers } = mount();
		section.update(
			baseState({
				shapeStyle: { fillMode: 'pattern', fillPatternPreset: 'cross' },
			} as unknown as Partial<InspectorState>),
		);
		const select = section.el.querySelector(
			'.pptxv-inspector-pattern pptx-ui-select',
		) as HTMLSelectElement;

		select.value = 'dotGrid';
		select.dispatchEvent(new Event('change'));

		expect(handlers.setShapeStyle).toHaveBeenCalledWith({
			fillMode: 'pattern',
			fillPatternPreset: 'dotGrid',
		});
	});
});

// B6 (A1/A2): "Recent colours" rows under the fill and stroke pickers.
describe('createFillSection recent-colours rows', () => {
	it('renders one recent-colours row per picker, hidden until there are entries', () => {
		const { section } = mount();
		section.update(baseState());

		const rows = section.el.querySelectorAll('[data-testid="pptx-color-recent"]');
		expect(rows).toHaveLength(2);
		for (const row of rows) {
			expect((row as HTMLElement).hidden).toBeTruthy();
		}

		section.update(baseState({ recentColors: ['#112233'] }));
		for (const row of section.el.querySelectorAll('[data-testid="pptx-color-recent"]')) {
			expect((row as HTMLElement).hidden).toBeFalsy();
		}
	});

	it('commits a fill pick through setShapeFill, and a stroke pick through setShapeStroke', () => {
		const { section, handlers } = mount();
		section.update(baseState({ recentColors: ['#112233'] }));

		const [fillRow, strokeRow] = Array.from(
			section.el.querySelectorAll<HTMLElement>('[data-testid="pptx-color-recent"]'),
		);
		recentSwatches(fillRow)[0]!.click();
		expect(handlers.setShapeFill).toHaveBeenCalledExactlyOnceWith('#112233');
		expect(handlers.setShapeStroke).not.toHaveBeenCalled();

		recentSwatches(strokeRow)[0]!.click();
		expect(handlers.setShapeStroke).toHaveBeenCalledExactlyOnceWith('#112233');
	});

	it('disables both rows when the shape cannot be formatted', () => {
		const { section } = mount();
		section.update(baseState({ canShape: false, recentColors: ['#112233'] }));

		for (const row of section.el.querySelectorAll('[data-testid="pptx-color-recent"]')) {
			expect(recentSwatches(row)[0]!.disabled).toBeTruthy();
		}
	});
});

describe('createFillSection theme colour swatch grid (W3-G2)', () => {
	it('renders one theme-swatch grid per picker, hidden until a theme is loaded', () => {
		const { section } = mount();
		section.update(baseState());

		const grids = themeGrids(section.el);
		expect(grids).toHaveLength(2);
		for (const grid of grids) {
			expect(grid.hidden).toBeTruthy();
		}

		section.update(baseState({ themeColorMap: OFFICE_THEME }));
		for (const grid of themeGrids(section.el)) {
			expect(grid.hidden).toBeFalsy();
		}
	});

	it('clicking a theme swatch commits both the hex and the ref, per picker', () => {
		const { section, handlers } = mount();
		section.update(baseState({ themeColorMap: OFFICE_THEME }));

		const [fillGrid, strokeGrid] = themeGrids(section.el);
		themeSwatch(fillGrid, 'Accent 2')!.click();
		expect(handlers.setShapeFill).toHaveBeenCalledExactlyOnceWith('#ed7d31', { scheme: 'accent2' });
		expect(handlers.setShapeStroke).not.toHaveBeenCalled();

		themeSwatch(strokeGrid, 'Accent 2')!.click();
		expect(handlers.setShapeStroke).toHaveBeenCalledExactlyOnceWith('#ed7d31', {
			scheme: 'accent2',
		});
	});

	it('disables both grids when the shape cannot be formatted', () => {
		const { section } = mount();
		section.update(baseState({ canShape: false, themeColorMap: OFFICE_THEME }));

		const grids = themeGrids(section.el);
		expect(grids).toHaveLength(2);
		for (const grid of grids) {
			const swatches = themeSwatches(grid);
			expect(swatches).toHaveLength(60);
			expect(swatches.every((swatch) => swatch.disabled)).toBeTruthy();
		}
	});
});

describe('createFillSection Fill and Line groups', () => {
	it('puts each picker with its own theme grid and recent colours under a named group', () => {
		const { section, handlers } = mount();
		section.update(baseState({ themeColorMap: OFFICE_THEME, recentColors: ['#112233'] }));

		// The theme grids are groups of their own ("Theme Colors"), inside these.
		const groups = Array.from(section.el.querySelectorAll<HTMLElement>(':scope > [role="group"]'));
		expect(groups.map((group) => group.getAttribute('aria-label'))).toStrictEqual([
			'pptx.inspector.fill',
			'pptx.inspector.line',
		]);
		for (const group of groups) {
			// The visible heading says which target the grid below it sets.
			expect(group.querySelector('.pptxv-inspector-row-label')?.textContent).toBe(
				group.getAttribute('aria-label'),
			);
			expect(group.querySelectorAll('input[type="color"]')).toHaveLength(1);
			expect(themeGrids(group)).toHaveLength(1);
			expect(group.querySelectorAll('[data-testid="pptx-color-recent"]')).toHaveLength(1);
		}
		// No grid or recent row is left loose in the section, outside a group.
		for (const child of Array.from(section.el.children)) {
			expect(child.matches('pptx-ui-color-grid, [data-testid="pptx-color-recent"]')).toBeFalsy();
		}

		const [fillGroup, lineGroup] = groups;
		themeSwatch(lineGroup, 'Accent 2')!.click();
		expect(handlers.setShapeStroke).toHaveBeenCalledExactlyOnceWith('#ed7d31', {
			scheme: 'accent2',
		});
		expect(handlers.setShapeFill).not.toHaveBeenCalled();
		recentSwatches(fillGroup)[0]!.click();
		expect(handlers.setShapeFill).toHaveBeenCalledExactlyOnceWith('#112233');
	});
});

describe('createFillSection gradient stop theme colour grid', () => {
	function gradientState(): Partial<InspectorState> {
		return {
			gradientEnabled: true,
			gradient: {
				type: 'linear',
				angle: 90,
				stops: [
					{ color: '#ff0000', position: 0 },
					{ color: '#ed7d31', position: 100, colorRef: { scheme: 'accent2' } },
				],
			},
			themeColorMap: OFFICE_THEME,
		};
	}

	it('renders one theme-swatch grid per gradient stop, hidden until a theme is loaded', () => {
		const { section } = mount();
		section.update(baseState({ ...gradientState(), themeColorMap: undefined }));
		const grids = section.el.querySelectorAll<HTMLElement>(
			'.pptxv-inspector-gradient-stops pptx-ui-color-grid',
		);
		expect(grids).toHaveLength(2);
		for (const grid of grids) {
			expect(grid.hidden).toBeTruthy();
		}

		section.update(baseState(gradientState()));
		for (const grid of section.el.querySelectorAll<HTMLElement>(
			'.pptxv-inspector-gradient-stops pptx-ui-color-grid',
		)) {
			expect(grid.hidden).toBeFalsy();
		}
	});

	it('clicking a theme swatch commits both hex and ref for the right stop index', () => {
		const { section, handlers } = mount();
		section.update(baseState(gradientState()));

		const grids = section.el.querySelectorAll<HTMLElement>(
			'.pptxv-inspector-gradient-stops pptx-ui-color-grid',
		);
		themeSwatch(grids[0], 'Accent 2')!.click();
		expect(handlers.updateGradientStop).toHaveBeenCalledExactlyOnceWith(0, {
			color: '#ed7d31',
			colorRef: { scheme: 'accent2' },
		});
	});

	it("highlights the stop's stored theme ref", () => {
		const { section } = mount();
		section.update(baseState(gradientState()));

		const grids = section.el.querySelectorAll<HTMLElement>(
			'.pptxv-inspector-gradient-stops pptx-ui-color-grid',
		);
		const swatch = themeSwatch(grids[1], 'Accent 2')!;
		expect(swatch.getAttribute('aria-checked')).toBe('true');
	});

	it("the native colour input clears the stop's ref", () => {
		const { section, handlers } = mount();
		section.update(baseState(gradientState()));

		const colorInputs = section.el.querySelectorAll<HTMLInputElement>(
			'.pptxv-inspector-gradient-stops input[type="color"]',
		);
		colorInputs[1].value = '#00ff00';
		colorInputs[1].dispatchEvent(new Event('input'));
		expect(handlers.updateGradientStop).toHaveBeenCalledExactlyOnceWith(1, {
			color: '#00ff00',
			colorRef: undefined,
		});
	});
});
