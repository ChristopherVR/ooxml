/**
 * W3-G2: the theme colour picker grid renders the deck's real theme palette (via `ooxml-ui/pptx`'s
 * `buildThemeColorSwatchGrid`) on the shared `pptx-ui-color-grid` instead of a hard-coded Office
 * palette, highlights the swatch matching the element's current ref or hex, and commits both the
 * resolved hex and the ref on click. Named `*.svelte.test.ts` so the props can be `$state` and
 * change after mount, the way an inspector field's selection does.
 */
import type { ThemeColorPickerCommit } from 'ooxml-ui/pptx';
import type { PptxThemeColorRef } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it } from 'vitest';

import { themeSwatch } from './theme-color-grid.test-support';
import ThemeColorSwatchGrid from './ThemeColorSwatchGrid.svelte';

const OFFICE_THEME = {
	dk1: '#000000',
	lt1: '#FFFFFF',
	dk2: '#44546A',
	lt2: '#E7E6E6',
	accent1: '#4472C4',
	accent2: '#ED7D31',
	accent3: '#A5A5A5',
	accent4: '#FFC000',
	accent5: '#5B9BD5',
	accent6: '#70AD47',
	bg1: '#FFFFFF',
	tx1: '#000000',
	bg2: '#E7E6E6',
	tx2: '#44546A',
};

interface GridProps {
	themeColorMap: Record<string, string> | undefined;
	selectedRef?: PptxThemeColorRef;
	selectedHex?: string;
	disabled?: boolean;
	onpick: (commit: ThemeColorPickerCommit) => void;
}

let cleanup: (() => void) | undefined;
afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

function render(props: GridProps): HTMLElement {
	const target = document.createElement('div');
	document.body.appendChild(target);
	const instance = mount(ThemeColorSwatchGrid, { target, props });
	flushSync();
	cleanup = () => {
		unmount(instance);
		target.remove();
	};
	return target;
}

const shadowOf = (target: HTMLElement) => target.querySelector('pptx-ui-color-grid')!.shadowRoot!;
const checkedTitles = (target: HTMLElement) =>
	[...shadowOf(target).querySelectorAll('button[aria-checked="true"]')].map((button) =>
		button.getAttribute('title'),
	);

describe('themeColorSwatchGrid', () => {
	it('renders nothing when no deck theme is loaded', () => {
		const target = render({ themeColorMap: undefined, onpick: () => {} });
		expect(target.querySelector('pptx-ui-color-grid')).toBeNull();
		expect(target.querySelector('button')).toBeNull();
	});

	it('renders a base + variant swatch per theme column on the shared grid', () => {
		const target = render({ themeColorMap: OFFICE_THEME, onpick: () => {} });
		// The swatches live in the shared element's shadow root, not in the binding's markup.
		expect(target.querySelector('button')).toBeNull();
		const shadow = shadowOf(target);
		// 10 columns x (1 base + 5 variants) = 60 swatch buttons, and no standard row.
		expect(shadow.querySelectorAll('button.swatch[data-source="theme"]')).toHaveLength(60);
		expect(shadow.querySelectorAll('button.swatch')).toHaveLength(60);
		// Only the Theme Colors heading: the inspector keeps its own standard and recent rows.
		const headings = [...shadow.querySelectorAll('.heading')].map((el) => el.textContent);
		expect(headings).toStrictEqual(['Theme Colors']);
		expect(target.querySelector('pptx-ui-color-grid')!.getAttribute('aria-label')).toBe(
			'Theme Colors',
		);
	});

	it('highlights the swatch matching the element ref, and follows it when the ref changes', () => {
		const props = $state<GridProps>({
			themeColorMap: OFFICE_THEME,
			selectedRef: { scheme: 'accent1', lumMod: 0.2, lumOff: 0.8 },
			onpick: () => {},
		});
		const target = render(props);
		expect(checkedTitles(target)).toStrictEqual(['Accent 1, Lighter 80%']);

		props.selectedRef = { scheme: 'accent2' };
		flushSync();
		expect(checkedTitles(target)).toStrictEqual(['Accent 2']);
	});

	it('highlights the swatch matching the hex when there is no ref', () => {
		const target = render({
			themeColorMap: OFFICE_THEME,
			selectedHex: '#ED7D31',
			onpick: () => {},
		});
		expect(checkedTitles(target)).toStrictEqual(['Accent 2']);
	});

	it('disables every swatch when the field is disabled', () => {
		const props = $state<GridProps>({
			themeColorMap: OFFICE_THEME,
			disabled: true,
			onpick: () => {},
		});
		const target = render(props);
		const swatches = () => [
			...shadowOf(target).querySelectorAll<HTMLButtonElement>('button.swatch'),
		];
		expect(swatches()).toHaveLength(60);
		expect(swatches().every((button) => button.disabled)).toBeTruthy();

		props.disabled = false;
		flushSync();
		expect(swatches().some((button) => button.disabled)).toBeFalsy();
	});

	it('commits both the hex and the ref on click', () => {
		const commits: ThemeColorPickerCommit[] = [];
		const target = render({ themeColorMap: OFFICE_THEME, onpick: (c) => commits.push(c) });
		const accent1Base = themeSwatch(target, 'Accent 1');
		expect(accent1Base).not.toBeNull();
		accent1Base!.click();
		const lighter = themeSwatch(target, 'Accent 1, Lighter 80%')!;
		lighter.click();
		expect(commits).toStrictEqual([
			{ hex: '#4472c4', ref: { scheme: 'accent1' } },
			{ hex: lighter.dataset.color, ref: { scheme: 'accent1', lumMod: 0.2, lumOff: 0.8 } },
		]);
	});
});
