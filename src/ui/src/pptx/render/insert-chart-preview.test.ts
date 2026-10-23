// @vitest-environment jsdom
import { describe, expect, it } from 'vitest';

import { parseSvgPreview } from '../../ribbon/safe-svg';
import { insertChartPreviewSvg } from './insert-chart-preview';

const SIZE = { width: 360, height: 220 };

describe('insertChartPreviewSvg', () => {
	it('draws the chart the dialog would insert as standalone SVG markup', () => {
		const svg = insertChartPreviewSvg(
			document,
			{ type: 'column', grouping: 'clustered', title: 'Sales' },
			SIZE,
			false,
		);
		const node = parseSvgPreview(document, svg);
		expect(node?.localName).toBe('svg');
		expect(svg).toContain('Sales');
		expect(node!.querySelectorAll('[data-chart-part="dataPoint"]').length).toBeGreaterThan(0);
	});

	it('leaves the title and legend out of a tile', () => {
		const svg = insertChartPreviewSvg(
			document,
			{ type: 'pie', grouping: 'clustered', title: 'Sales' },
			SIZE,
			true,
		);
		expect(svg).not.toContain('Sales');
		expect(svg).not.toContain('Chart Title');
	});
});
