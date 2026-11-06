import { dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

import { describe, expect, it } from 'vitest';

import { componentSource } from './component-source.test-support';

/**
 * PowerPoint's Normal view paints no transport at rest and a control bar under
 * the shape while it is selected or hovered. Angular mounts the shared bar on
 * the element box and threads the canvas selection down to it (this package has
 * no TestBed, so this is asserted against the authored templates and class; the
 * bar's behaviour itself is covered by the shared helper's tests).
 */
describe('the authoring-canvas media control bar', () => {
	const dir = dirname(fileURLToPath(import.meta.url));
	const read = (name: string) => componentSource(dir, name);

	it('mounts the shared bar for the authoring surface only, with the selection as input', () => {
		const source = read('media-renderer.component.ts');
		expect(source).toContain('mediaControlBarEnabled(this.surface())');
		expect(source).toContain('attachMediaControlBar(host, el, {');
		expect(source).toContain('readonly selected = input<boolean>(false);');
		expect(source).toContain('controlBar?.setSelected(selected)');
	});

	it('threads the canvas selection from the slide canvas to the media renderer', () => {
		expect(read('slide-canvas.component.html')).toContain(
			'[selected]="selectedIds().includes(element.id)"',
		);
		expect(read('element-renderer.component.html')).toContain('[selected]="selected()"');
		expect(read('element-renderer-graphics.component.html')).toContain('[selected]="selected()"');
	});
});
