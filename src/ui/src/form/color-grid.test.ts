import { buildThemePalette } from 'ooxml-core/color';
import { afterEach, beforeAll, describe, expect, it, vi } from 'vitest';
import type { OfficeColorPick, OfficeUiColorGrid } from '../controls';
import { registerOfficeUi } from '../index';

beforeAll(() => registerOfficeUi());
afterEach(() => document.body.replaceChildren());

function make(setup: (grid: OfficeUiColorGrid) => void = () => {}): OfficeUiColorGrid {
	const grid = document.createElement('office-ui-color-grid') as OfficeUiColorGrid;
	setup(grid);
	document.body.append(grid);
	return grid;
}
const buttons = (grid: OfficeUiColorGrid, selector = 'button') => [
	...grid.shadowRoot!.querySelectorAll<HTMLButtonElement>(selector),
];
const key = (target: Element, name: string) =>
	target.dispatchEvent(
		new KeyboardEvent('keydown', { key: name, bubbles: true, composed: true, cancelable: true }),
	);

describe('office-ui-color-grid', () => {
	it('shows the Office theme grid and the standard colours by default', () => {
		const grid = make();
		expect(grid.getAttribute('role')).toBe('group');
		expect(grid.getAttribute('aria-label')).toBe('Colors');
		expect(
			[...grid.shadowRoot!.querySelectorAll('.heading')].map((heading) => heading.textContent),
		).toEqual(['Theme Colors', 'Standard Colors']);
		// Ten base colours, five rows of variants, ten standard colours; no commands.
		expect(buttons(grid, '[data-source="theme"]')).toHaveLength(60);
		expect(buttons(grid, '[data-source="standard"]')).toHaveLength(10);
		expect(buttons(grid, '.command')).toHaveLength(0);
		const accent = buttons(grid, '[data-color="#4472c4"]')[0]!;
		expect(accent.title).toBe('Accent 1');
		expect(buttons(grid, '[data-color="#d9e2f3"]')[0]!.getAttribute('aria-label')).toBe(
			'Accent 1, Lighter 80%',
		);
		expect(buttons(grid, '[data-color="#ff0000"]')[0]!.title).toBe('Red');
	});

	it('builds the theme grid from the given colours and marks the current one', () => {
		const grid = make((el) => {
			el.themeColors = ['#ffffff', '#000000', '#eeeeee', '#222222', '#112233'];
			el.value = '#112233';
			el.label = 'Fill colours';
		});
		expect(grid.getAttribute('aria-label')).toBe('Fill colours');
		const chosen = buttons(grid, '[aria-checked="true"]');
		expect(chosen.map((button) => button.dataset.color)).toEqual(['#112233']);
		expect(chosen[0]!.dataset.column).toBe('4');
		// A column the theme does not give keeps the Office colour.
		expect(buttons(grid, '[data-color="#70ad47"]')).toHaveLength(1);
	});

	it('emits the picked colour with its source and name', () => {
		const grid = make((el) => {
			el.noneLabel = 'No Fill';
			el.automaticLabel = 'Automatic';
			el.moreLabel = 'More Colors...';
			el.recentColors = ['ABCDEF', 'not a colour', '#abcdef'];
		});
		const picked: OfficeColorPick[] = [];
		const more = vi.fn();
		document.body.addEventListener('office-color-pick', (event) =>
			picked.push((event as CustomEvent<OfficeColorPick>).detail),
		);
		document.body.addEventListener('office-color-more', more);
		expect(buttons(grid, '.command').map((button) => button.textContent!.trim())).toEqual([
			'Automatic',
			'No Fill',
			'More Colors...',
		]);
		// Recent colours are normalised; the bad entry is dropped.
		expect(buttons(grid, '[data-source="recent"]').map((button) => button.dataset.color)).toEqual([
			'#abcdef',
			'#abcdef',
		]);
		buttons(grid, '[data-command="automatic"]')[0]!.click();
		buttons(grid, '[data-command="none"]')[0]!.click();
		buttons(grid, '[data-color="#ed7d31"]')[0]!.click();
		buttons(grid, '[data-color="#7030a0"]')[0]!.click();
		buttons(grid, '[data-source="recent"]')[0]!.click();
		buttons(grid, '[data-command="more"]')[0]!.click();
		expect(picked).toEqual([
			{ color: 'automatic', source: 'automatic', label: 'Automatic' },
			{ color: 'none', source: 'none', label: 'No Fill' },
			{ color: '#ed7d31', source: 'theme', label: 'Accent 2', theme: { column: 5 } },
			{ color: '#7030a0', source: 'standard', label: 'Purple' },
			{ color: '#abcdef', source: 'recent', label: '#ABCDEF' },
		]);
		expect(more).toHaveBeenCalledTimes(1);
	});

	it('marks No Fill as the current choice', () => {
		const grid = make((el) => {
			el.noneLabel = 'No Fill';
			el.value = 'none';
		});
		expect(buttons(grid, '[data-command="none"]')[0]!.getAttribute('aria-checked')).toBe('true');
		expect(buttons(grid, '.swatch[aria-checked="true"]')).toHaveLength(0);
	});

	it('moves with the arrow keys, Home and End, and keeps the keys from a host menu', () => {
		const grid = make((el) => {
			el.noneLabel = 'No Fill';
			el.moreLabel = 'More Colors...';
		});
		const host = vi.fn();
		document.body.addEventListener('keydown', host);
		const root = grid.shadowRoot!;
		const active = () => root.activeElement as HTMLButtonElement;
		grid.focus();
		expect(active().dataset.command).toBe('none');
		key(active(), 'ArrowDown');
		expect(active().dataset.color).toBe('#ffffff');
		key(active(), 'ArrowRight');
		key(active(), 'ArrowRight');
		expect(active().title).toBe('Background 2');
		key(active(), 'ArrowDown');
		expect(active().title).toBe('Background 2, Darker 10%');
		key(active(), 'ArrowUp');
		key(active(), 'ArrowLeft');
		expect(active().title).toBe('Text 1');
		key(active(), 'End');
		expect(active().dataset.command).toBe('more');
		key(active(), 'ArrowUp');
		// A one-item row lands on the nearest column of the row above.
		expect(active().title).toBe('Dark Red');
		key(active(), 'Home');
		expect(active().dataset.command).toBe('none');
		expect(host).not.toHaveBeenCalled();
		// Past the first row, and for other keys, the key is the host's.
		key(active(), 'ArrowUp');
		expect(active().dataset.command).toBe('none');
		key(active(), 'Escape');
		expect(host).toHaveBeenCalledTimes(2);
	});

	it('takes product column names and an extra row, and reports where a swatch came from', () => {
		const grid = make((el) => {
			el.themeNames = ['White', 'Black', 'Light', 'Dark'];
			el.themeColors = ['#ffffff', '#000000', '#ffffff', '#000000'];
			el.extraColors = [
				{ hex: '#123456', label: 'Variant Accent 1' },
				{ hex: '#654321', label: 'Variant Accent 2' },
			];
		});
		const picked: OfficeColorPick[] = [];
		grid.addEventListener('office-color-pick', (event) =>
			picked.push((event as CustomEvent<OfficeColorPick>).detail),
		);
		// Missing names fall back to the Office ones.
		expect(buttons(grid, '[data-source="theme"]')[2]!.title).toBe('Light');
		expect(buttons(grid, '[data-source="theme"]')[4]!.title).toBe('Accent 1');
		// The extra row sits between the variants and Standard Colors, with no heading of its own.
		const extra = buttons(grid, '[data-source="extra"]');
		expect(extra.map((button) => button.title)).toEqual(['Variant Accent 1', 'Variant Accent 2']);
		expect(
			[...grid.shadowRoot!.querySelectorAll('.heading')].map((heading) => heading.textContent),
		).toEqual(['Theme Colors', 'Standard Colors']);
		buttons(grid, '[aria-label="Dark, Lighter 50%"]')[0]!.click();
		extra[1]!.click();
		expect(picked).toEqual([
			{
				color: '#7f7f7f',
				source: 'theme',
				label: 'Dark, Lighter 50%',
				theme: { column: 3, variant: { kind: 'lighter', percent: 50 } },
			},
			{ color: '#654321', source: 'extra', label: 'Variant Accent 2', index: 1 },
		]);
	});

	it('shows a product palette as it is given, with the columns it has', () => {
		// Two columns whose swatches the product computed itself (rounded, not truncated).
		const palette = buildThemePalette()
			.slice(4, 6)
			.map((column, index) => ({
				...column,
				name: `Column ${index}`,
				variants: column.variants.map((variant, row) => ({
					...variant,
					hex: `#0${index}${row}0${index}${row}`,
					label: `Column ${index} variant ${row}`,
				})),
			}));
		const heard = vi.fn<(pick: OfficeColorPick) => void>();
		const grid = make((g) => {
			g.palette = palette;
			g.variantLabel = () => 'not used';
		});
		grid.addEventListener('office-color-pick', (event) =>
			heard((event as CustomEvent<OfficeColorPick>).detail),
		);
		const theme = buttons(grid, '[data-source="theme"]');
		expect(theme).toHaveLength(12);
		// Each theme row lays out the palette's two columns, not ten with eight empty.
		const rows = [...grid.shadowRoot!.querySelectorAll<HTMLElement>('[data-cells="theme"]')];
		expect(rows.map((row) => row.style.getPropertyValue('--_columns'))).toStrictEqual(
			Array(6).fill('2'),
		);
		expect(
			[...grid.shadowRoot!.querySelectorAll<HTMLElement>('[data-cells="standard"]')].map((row) =>
				row.style.getPropertyValue('--_columns'),
			),
		).toStrictEqual(['10']);
		expect(theme.map((button) => button.dataset.color)).toContain('#014014');
		const variant = buttons(grid, '[data-color="#014014"]')[0]!;
		expect(variant.getAttribute('aria-label')).toBe('Column 1 variant 4');
		variant.click();
		expect(heard).toHaveBeenCalledWith({
			color: '#014014',
			source: 'theme',
			label: 'Column 1 variant 4',
			theme: { column: 1, variant: palette[1]!.variants[4]!.variant },
		});
	});

	it('keeps a short column of a product palette in place instead of failing to draw', () => {
		// The second column resolves only four variants: its fifth slot stays empty.
		const [first, second] = buildThemePalette().slice(4, 6);
		const palette = [first!, { ...second!, variants: second!.variants.slice(0, 4) }];
		const heard = vi.fn<(pick: OfficeColorPick) => void>();
		const grid = make((g) => {
			g.palette = palette;
		});
		grid.addEventListener('office-color-pick', (event) =>
			heard((event as CustomEvent<OfficeColorPick>).detail),
		);
		expect(buttons(grid, '[data-source="theme"]')).toHaveLength(11);
		const lastRow = [...grid.shadowRoot!.querySelectorAll('[data-cells="theme"]')].at(-1)!;
		expect([...lastRow.children].map((cell) => cell.localName)).toStrictEqual(['button', 'span']);
		buttons(grid, `[title="${second!.variants[3]!.label}"]`)[0]!.click();
		expect(heard.mock.calls[0]![0].theme).toStrictEqual({
			column: 1,
			variant: second!.variants[3]!.variant,
		});
	});

	it('marks the swatch at selectedTheme and no swatch by value', () => {
		const grid = make((g) => {
			g.value = '#ff0000';
			g.selectedTheme = { column: 4, variant: { kind: 'lighter', percent: 80 } };
		});
		const marked = buttons(grid, '[aria-checked="true"]');
		expect(marked).toHaveLength(1);
		expect(marked[0]!.getAttribute('aria-label')).toBe('Accent 1, Lighter 80%');
		grid.selectedTheme = { column: 4 };
		return grid.updateComplete.then(() => {
			expect(buttons(grid, '[aria-checked="true"]').map((b) => b.title)).toEqual(['Accent 1']);
		});
	});

	it('leaves out Standard Colors when the product has none', () => {
		const grid = make((g) => {
			g.standardColors = [];
		});
		expect(
			[...grid.shadowRoot!.querySelectorAll('.heading')].map((heading) => heading.textContent),
		).toEqual(['Theme Colors']);
		expect(buttons(grid, '[data-source="standard"]')).toHaveLength(0);
	});

	it('shows only Recent Colors for an empty palette, each named by recentLabel', () => {
		const grid = make((el) => {
			el.palette = [];
			el.standardColors = [];
			el.recentColors = ['#FF0000', '#0000ff'];
			el.recentHeading = 'Zuletzt verwendete Farben';
			el.recentLabel = (hex) => `Zuletzt ${hex}`;
		});
		expect(
			[...grid.shadowRoot!.querySelectorAll('.heading')].map((heading) => heading.textContent),
		).toEqual(['Zuletzt verwendete Farben']);
		expect(
			buttons(grid).map((button) => [
				button.dataset.source,
				button.dataset.color,
				button.getAttribute('aria-label'),
			]),
		).toEqual([
			['recent', '#ff0000', 'Zuletzt #FF0000'],
			['recent', '#0000ff', 'Zuletzt #0000FF'],
		]);
		const picked = vi.fn();
		grid.addEventListener('office-color-pick', (event) => picked(event.detail));
		buttons(grid)[1]!.click();
		expect(picked).toHaveBeenCalledWith({
			color: '#0000ff',
			source: 'recent',
			label: 'Zuletzt #0000FF',
		});
	});

	it('does nothing while disabled', () => {
		const grid = make((el) => {
			el.disabled = true;
			el.moreLabel = 'More Colors...';
		});
		const heard = vi.fn();
		grid.addEventListener('office-color-pick', heard);
		grid.addEventListener('office-color-more', heard);
		expect(buttons(grid).every((button) => button.disabled)).toBe(true);
		buttons(grid, '.swatch')[0]!.click();
		expect(heard).not.toHaveBeenCalled();
	});
});
