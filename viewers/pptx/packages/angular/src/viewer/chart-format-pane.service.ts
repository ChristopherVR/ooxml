/**
 * chart-format-pane.service.ts: which chart element the Format pane and Chart
 * Format > Current Selection target, and whether the pane is open. Shared by
 * the chart canvas view (clicks retarget, a double-click opens), the ribbon
 * strip, the context menu and the pane host, without threading state through
 * the canvas layers.
 *
 * The target belongs to one chart element: reading it for another element
 * yields the chart area, so selecting a different chart starts over.
 *
 * Provided at the viewer level; consumers inject it `optional`.
 */
import { Injectable, signal } from '@angular/core';

@Injectable()
export class ChartFormatPaneService {
	private readonly state = signal<{ elementId: string | null; target: string }>({
		elementId: null,
		target: 'chartArea',
	});
	/** Whether the Format pane is open. */
	readonly open = signal(false);

	/** The current target of `elementId` (the chart area for any other element). */
	targetFor(elementId: string | null | undefined): string {
		const current = this.state();
		return elementId && current.elementId === elementId ? current.target : 'chartArea';
	}

	/** Make `target` the current element of chart `elementId`. */
	setTarget(elementId: string, target: string): void {
		this.state.set({ elementId, target: target || 'chartArea' });
	}

	/** Open the pane on `target` of chart `elementId`. */
	openOn(elementId: string, target: string): void {
		this.setTarget(elementId, target);
		this.open.set(true);
	}

	close(): void {
		this.open.set(false);
	}
}
