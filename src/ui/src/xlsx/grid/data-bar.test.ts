// @vitest-environment jsdom
import { readFileSync } from 'node:fs';
import { resolve } from 'node:path';
import JSZip from 'jszip';
import { createWorkbook, loadXlsx, saveXlsx } from 'ooxml-core/xlsx';
import { afterEach, describe, expect, it } from 'vitest';
import { mountGrid } from './index.js';
import { createTestContext } from './test-context.js';

const native = JSON.parse(
	readFileSync(
		resolve(process.cwd(), '../core/xlsx/edit/__fixtures__/excel-databar-clipboard.json'),
		'utf8',
	),
) as {
	cases: { variant: number; before: string }[];
};
let dispose: (() => void) | undefined;
afterEach(() => {
	dispose?.();
	dispose = undefined;
	document.body.replaceChildren();
});

describe('Data-bar grid rendering', () => {
	it('paints native solid RTL bars with distinct negative colors and borders, then repaints changed settings', async () => {
		const zip = await JSZip.loadAsync(await saveXlsx(createWorkbook()));
		zip.file('xl/worksheets/sheet1.xml', native.cases.find((c) => c.variant === 1)!.before);
		const workbook = await loadXlsx(await zip.generateAsync({ type: 'uint8array' }));
		const ctx = createTestContext(workbook);
		const container = document.createElement('div');
		ctx.root.append(container);
		dispose = mountGrid(ctx, container);
		const bars = [...ctx.root.querySelectorAll<HTMLElement>('.xg-db')];
		const negative = bars.find((bar) => bar.style.backgroundColor === 'rgb(0, 0, 255)')!;
		const positive = bars.filter((bar) => bar.style.backgroundColor === 'rgb(0, 255, 0)').at(-1)!;
		expect(negative).toBeDefined();
		expect(positive).toBeDefined();
		expect(negative.style.borderColor).toBe('rgb(255, 255, 0)');
		expect(positive.style.borderColor).toBe('rgb(255, 0, 0)');
		const cellWidth = Number.parseFloat(negative.parentElement!.style.width);
		expect(Number.parseFloat(negative.style.left)).toBe(2 + (cellWidth - 4) / 2);
		expect(negative.style.right).toBe('auto');
		const axis = negative.parentElement!.querySelector<HTMLElement>('.xg-db-axis')!;
		expect(axis.style.left).toBe(negative.style.left);
		expect(axis.style.borderLeftColor).toBe('rgb(255, 0, 255)');
		expect(axis.style.borderLeftStyle).toBe('dashed');
		expect(positive.style.left).toBe('2px');
		expect(negative.style.backgroundImage).toBe('none');
		expect(negative.parentElement?.querySelector('.xg-tx')).toBeNull();

		const format = workbook.sheets[0]!.conditionalFormats[0]!;
		const rule = format.rules[0]!;
		if (rule.type !== 'dataBar') throw new Error('Expected a data bar');
		delete rule.extensionXml;
		delete rule.showValue;
		ctx.notify({ kind: 'format' });
		await new Promise((resolve) => setTimeout(resolve, 40));
		const repainted = [...ctx.root.querySelectorAll<HTMLElement>('.xg-db')];
		expect(repainted.length).toBe(bars.length);
		for (const bar of repainted) {
			expect(bar.style.left).toBe('2px');
			expect(bar.style.right).toBe('auto');
			expect(bar.style.borderStyle).toBe('none');
			expect(bar.style.backgroundImage).toContain('linear-gradient(90deg');
		}
		expect(ctx.root.querySelector('.xg-c .xg-tx')?.textContent).toBe('-20');
	});
});
