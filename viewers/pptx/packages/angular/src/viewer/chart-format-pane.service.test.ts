/**
 * The chart Format pane state: one target per chart element, defaulting to
 * the chart area, and the open flag the host, the ribbon strip, the context
 * menu and a chart double-click share.
 */
import { buildContextMenuEntries } from 'ooxml-ui/pptx';
import { describe, expect, it } from 'vitest';

import { ChartFormatPaneService } from './chart-format-pane.service';

describe('ChartFormatPaneService', () => {
	it('scopes the target to its chart and starts every other chart at the chart area', () => {
		const pane = new ChartFormatPaneService();
		expect(pane.targetFor('c1')).toBe('chartArea');
		pane.setTarget('c1', 'series:1');
		expect(pane.targetFor('c1')).toBe('series:1');
		expect(pane.targetFor('c2')).toBe('chartArea');
		expect(pane.targetFor(null)).toBe('chartArea');
		expect(pane.open()).toBe(false);
	});

	it('opens on a target and closes', () => {
		const pane = new ChartFormatPaneService();
		pane.openOn('c1', 'axis:0');
		expect(pane.open()).toBe(true);
		expect(pane.targetFor('c1')).toBe('axis:0');
		pane.close();
		expect(pane.open()).toBe(false);
		pane.setTarget('c1', '');
		expect(pane.targetFor('c1')).toBe('chartArea');
	});

	it('names the context-menu entry after the stored target', () => {
		const pane = new ChartFormatPaneService();
		pane.setTarget('c1', 'legend');
		const entry = buildContextMenuEntries({
			elementType: 'chart',
			chartFormatTarget: pane.targetFor('c1'),
		}).find((e) => e.id === 'format-chart-element');
		expect(entry?.labelKey).toBe('pptx.contextMenu.formatLegend');
	});
});
