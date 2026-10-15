import { chartFormatTargetFromEvent } from 'ooxml-ui/pptx';
import type { ChartFormatPointerEvent } from 'ooxml-ui/pptx';
import type { PptxElement } from 'pptx-viewer-core';

import type { ChartFormatState } from '../editor/chart-format-state.svelte';

/**
 * A press on a selected chart makes the element under the pointer the chart
 * Format target (Chart Format > Current Selection follows it, as in
 * PowerPoint). The hit-test itself is shared (`chartFormatTargetFromEvent`).
 */
export function chartFormatPointerDown(
	event: ChartFormatPointerEvent,
	root: HTMLElement | null,
	element: PptxElement,
	state: ChartFormatState | undefined,
): void {
	if (state && root && element.type === 'chart') {
		state.setTarget(element.id, chartFormatTargetFromEvent(event, root, element));
	}
}

/**
 * A double-click on a selected chart opens the Format pane on the element
 * under the pointer. The title is left to the inline title editor, so this
 * returns false for it; otherwise it handles the event and returns true.
 */
export function chartFormatDoubleClick(
	event: ChartFormatPointerEvent & { stopPropagation(): void },
	root: HTMLElement | null,
	element: PptxElement,
	state: ChartFormatState | undefined,
): boolean {
	if (!state || !root || element.type !== 'chart') {
		return false;
	}
	const target = chartFormatTargetFromEvent(event, root, element);
	if (target === 'title') {
		return false;
	}
	event.stopPropagation();
	state.openPane(element.id, target);
	return true;
}
