/**
 * Touch long-press and double-tap wiring (issue #290). iOS Safari fires no
 * `contextmenu` on a long-press and no reliable `dblclick`, so the viewer root
 * installs the shared `installTouchPress` and each control that relies on those
 * events opts in with `data-pptx-touch-press`. The Angular suite does not mount
 * templates, so this pins the opt-in on the same tag as the handler it serves.
 */
import { readFileSync } from 'node:fs';
import path from 'node:path';

import { describe, expect, it } from 'vitest';

const read = (file: string): string => readFileSync(path.join(__dirname, file), 'utf8');

/** The opening tag that contains `marker`. */
function tagAround(source: string, marker: string): string {
	const at = source.indexOf(marker);
	expect(at, marker).toBeGreaterThan(-1);
	const start = source.lastIndexOf('<', at);
	const end = source.indexOf('>', at);
	return source.slice(start, end);
}

describe('touch press opt-in', () => {
	it.each([
		['slides-panel.component.html', '(contextmenu)="onThumbContextMenu', 'menu'],
		['slides-panel.component.html', 'onSectionContextMenu($event', 'menu dblclick'],
		['slide-sorter-overlay.component.html', '(dblclick)="select.emit(i)"', 'menu dblclick'],
		['slide-canvas.component.html', '(contextmenu)="onContextMenu($event)"', 'menu'],
		['presentation-overlay.component.html', '(contextmenu)="onStageContextMenu', 'menu'],
	])('%s opts %s into %s', (file, marker, actions) => {
		expect(tagAround(read(file), marker)).toContain(`data-pptx-touch-press="${actions}"`);
	});

	it('opts an editable chart into double-tap for its part Format pane', () => {
		expect(
			tagAround(read('chart-element-view.component.html'), '(dblclick)="onDblClick'),
		).toContain(`[attr.data-pptx-touch-press]="canEdit() ? 'dblclick' : null"`);
	});

	it('installs the shared handler on the viewer root instead of a selection-only long-press', () => {
		const service = read('viewer-touch-gestures.service.ts');
		expect(service).toMatch(/installTouchPress\(el\)/);
		expect(service).not.toMatch(/onLongPress/);
		expect(read('power-point-viewer.component.ts')).toMatch(
			/setupTouchPress\(\(\) => this\.viewerRoot\(\)\?\.nativeElement\)/,
		);
	});
});
