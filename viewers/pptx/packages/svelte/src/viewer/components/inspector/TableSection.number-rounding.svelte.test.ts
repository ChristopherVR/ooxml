import type { PptxElement } from 'pptx-viewer-core';
import { flushSync, mount, unmount } from 'svelte';
import { afterEach, describe, expect, it } from 'vitest';

import { EditorState } from '../../editor/editor-state.svelte';
import TableSection from './TableSection.svelte';

let cleanup: (() => void) | undefined;

afterEach(() => {
	cleanup?.();
	cleanup = undefined;
});

// 1.52pt borders, 4pt margins and a 29.2pt row, read back from EMU as px.
const BORDER = 19304 / 9525;
const MARGIN = 50800 / 9525;
const ROW = 370840 / 9525;

function tableEl(): PptxElement {
	return {
		type: 'table',
		id: 'tbl1',
		x: 0,
		y: 0,
		width: 200,
		height: 100,
		tableData: {
			rows: [
				{
					height: ROW,
					cells: [
						{
							text: 'A',
							style: {
								borderTopWidth: BORDER,
								borderBottomWidth: BORDER,
								borderLeftWidth: BORDER,
								borderRightWidth: BORDER,
								marginTop: MARGIN,
								marginBottom: MARGIN,
								marginLeft: MARGIN,
								marginRight: MARGIN,
							},
						},
					],
				},
			],
			columnWidths: [1],
		},
	} as PptxElement;
}

describe('table inspector number display', () => {
	it('shows EMU-derived widths, margins and row heights with two decimals at most', () => {
		const el = tableEl();
		const editor = new EditorState({ getCurrent: () => 0, getHandler: () => null });
		editor.editable = true;
		editor.setSlides([{ id: 's1', rId: 'rId1', slideNumber: 1, elements: [el] }]);
		editor.select(el.id);
		const target = document.createElement('div');
		document.body.appendChild(target);
		const instance = mount(TableSection, { target, props: { editor, el } });
		flushSync();
		cleanup = () => {
			unmount(instance);
			target.remove();
		};

		const values = Array.from(
			target.querySelectorAll<HTMLInputElement>('input[type="number"]'),
		).map((input) => input.value);
		expect(values.filter((value) => value === '2.03')).toHaveLength(4);
		expect(values.filter((value) => value === '5.33')).toHaveLength(4);
		expect(values).toContain('38.93');
		expect(values.filter((value) => /\.\d{3,}/u.test(value))).toStrictEqual([]);
	});
});
