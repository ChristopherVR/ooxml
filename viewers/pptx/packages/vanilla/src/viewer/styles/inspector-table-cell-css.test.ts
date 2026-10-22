import { afterEach, describe, expect, it } from 'vitest';

import { createTableCellFillControls } from '../ui/inspector/table-cell-fill-controls';
import { createTableCellFormatting } from '../ui/inspector/table-cell-formatting';
import { createTableStructureControls } from '../ui/inspector/table-structure-controls';
import type { InspectorHandlers, InspectorState } from '../ui/inspector/types';
import { createThemeColorSwatchGrid } from '../ui/theme-color-swatch-grid';
import { buildViewerCss } from './css';

const THEME: Record<string, string> = {
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
};

const STATE = {
	isTable: true,
	selectedTableCell: { row: 0, column: 0 },
	selectedTableCells: [],
	tableColumnWidths: [0.5, 0.5],
	tableRowHeights: [32, 32],
	themeColorMap: THEME,
} as unknown as InspectorState;

const HANDLERS = {} as InspectorHandlers;
const identity = (key: string): string => key;

const mounted: Element[] = [];

/** Put the viewer stylesheet and `nodes` in the document; `afterEach` takes them out again. */
function mount(...nodes: Element[]): void {
	const style = document.createElement('style');
	style.textContent = buildViewerCss();
	document.head.appendChild(style);
	for (const node of nodes) {
		document.body.appendChild(node);
	}
	mounted.push(style, ...nodes);
}

afterEach(() => {
	for (const node of mounted.splice(0)) {
		node.remove();
	}
});

/** Whether `css` has a selector naming `.name` itself, not just a longer class it prefixes. */
function hasRuleFor(css: string, name: string): boolean {
	return new RegExp(`\\.${name}(?![\\w-])`, 'u').test(css);
}

/** Every class the given subtree renders, so each can be checked for a rule. */
function renderedClasses(root: HTMLElement): string[] {
	const classes = new Set<string>();
	for (const node of [root, ...root.querySelectorAll<HTMLElement>('[class]')]) {
		for (const name of node.classList) {
			if (name.startsWith('pptxv-')) {
				classes.add(name);
			}
		}
	}
	return [...classes];
}

describe('table cell inspector styles', () => {
	it('has a rule for every class the cell panels and the theme grid render', () => {
		const css = buildViewerCss();
		const cell = createTableCellFormatting(document, identity, HANDLERS);
		const fill = createTableCellFillControls(document, identity, HANDLERS);
		const structure = createTableStructureControls(document, identity, HANDLERS);
		for (const panel of [cell, fill, structure]) {
			panel.update(STATE);
		}

		const classes = [cell.el, fill.el, structure.el].flatMap(renderedClasses);
		// The theme grids are among what is checked: the shared element, under the host's class.
		expect(cell.el.querySelectorAll('pptx-ui-color-grid.pptxv-theme-swatch-grid')).toHaveLength(2);
		expect(classes).toContain('pptxv-theme-swatch-grid');
		const missing = classes.filter((name) => !hasRuleFor(css, name));
		expect(missing).toStrictEqual([]);
	});

	it('shows the theme grid as the shared element, spaced under its field', () => {
		const grid = createThemeColorSwatchGrid(document, identity, () => {});
		grid.setThemeColorMap(THEME);
		mount(grid.el);

		// The shared grid draws and sizes its own swatches; the stylesheet only spaces it.
		expect(grid.el.localName).toBe('pptx-ui-color-grid');
		expect(grid.el.shadowRoot!.querySelectorAll('button.swatch')).toHaveLength(60);
		expect(getComputedStyle(grid.el).marginTop).toBe('4px');
		expect(getComputedStyle(grid.el).display).not.toBe('none');

		grid.setThemeColorMap(undefined);
		expect(getComputedStyle(grid.el).display).toBe('none');
	});

	it('lays the cell fields out as rows instead of one run of inline text', () => {
		const cell = createTableCellFormatting(document, identity, HANDLERS);
		cell.update(STATE);
		mount(cell.el);

		expect(getComputedStyle(cell.el).display).toBe('grid');
		expect(getComputedStyle(cell.el.querySelector('label')!).display).toBe('flex');
		// Each theme grid is a row of its own, straight under its colour field.
		for (const grid of cell.el.querySelectorAll('pptx-ui-color-grid')) {
			expect(grid.parentElement).toBe(cell.el);
			expect(grid.previousElementSibling?.querySelector('input[type="color"]')).not.toBeNull();
		}
	});
});
