// @vitest-environment happy-dom
import { afterEach, describe, expect, it } from 'vitest';

import {
	settleTableBackgrounds,
	TABLE_BG_SPAN_VAR,
	tableBackgroundSlice,
} from './table-style-bg-settle';

describe('tableBackgroundSlice', () => {
	it('shows the part of a table-sized gradient under the cell', () => {
		expect(
			tableBackgroundSlice(
				{ left: 50, top: 80.5, width: 100, height: 40 },
				{ width: 300, height: 200 },
			),
		).toStrictEqual({ size: '300px 200px', position: '-50px -80.5px' });
	});

	it('keeps layers above the table gradient on their own box', () => {
		expect(
			tableBackgroundSlice(
				{ left: 0, top: 10, width: 10, height: 10 },
				{ width: 30, height: 20 },
				2,
			),
		).toStrictEqual({ size: 'auto, 30px 20px', position: '0 0, 0px -10px' });
	});
});

describe('settleTableBackgrounds', () => {
	afterEach(() => {
		document.body.replaceChildren();
	});

	function box(element: HTMLElement, values: Record<string, number>): void {
		for (const [key, value] of Object.entries(values)) {
			Object.defineProperty(element, key, { value, configurable: true });
		}
	}

	it('places a cell from the rows as laid out, not as declared', () => {
		// The first row was declared 20px but its text grew it to 50px: the second
		// row's slice must start at 50, where the declared heights would say 20.
		document.body.innerHTML = `<table><tbody>
			<tr><td id="a" style="${TABLE_BG_SPAN_VAR}: 1"></td></tr>
			<tr><td id="b" style="${TABLE_BG_SPAN_VAR}: 1"></td></tr>
		</tbody></table>`;
		box(document.querySelector('table') as HTMLElement, { offsetWidth: 100, offsetHeight: 80 });
		box(document.getElementById('a') as HTMLElement, {
			offsetLeft: 0,
			offsetTop: 0,
			offsetWidth: 100,
			offsetHeight: 50,
		});
		box(document.getElementById('b') as HTMLElement, {
			offsetLeft: 0,
			offsetTop: 50,
			offsetWidth: 100,
			offsetHeight: 30,
		});
		expect(settleTableBackgrounds()).toBe(2);
		const b = document.getElementById('b') as HTMLElement;
		expect(b.style.backgroundSize).toBe('100px 80px');
		expect(b.style.backgroundPosition).toBe('0px -50px');
	});
});
