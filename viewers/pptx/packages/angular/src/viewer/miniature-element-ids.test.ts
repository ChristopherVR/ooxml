/**
 * Only ONE node in the document may answer to a given `data-element-id`.
 *
 * Angular reuses the live `SlideCanvasComponent` for its miniatures and static
 * stages, unlike React, whose thumbnails go through a separate
 * `StaticElementRenderer` that stamps no id. The thumbnail rail once painted
 * every element of EVERY slide with its real id, so a framework-neutral
 * `[data-element-id]` query answered from a slide that was not on screen
 * (`e2e/custom-shows.spec.ts` read "Alpha Slide" out of the rail on every step
 * of a running show), and the presenter previews and reading view still
 * duplicated the editing canvas's ids after the rail was fixed.
 *
 * So no call site decides: the canvas derives it from the shared
 * `stageKeepsElementIds` rule (`static-stage-element-ids.test.ts` proves it
 * by rendering). These source guards keep the call sites on that rule.
 */
import { readdirSync, readFileSync } from 'node:fs';
import path from 'node:path';

import { describe, expect, it } from 'vitest';

import { componentSource } from './component-source.test-support';

const VIEWER_DIR = import.meta.dirname;

/** Authored viewer sources (components, templates), tests excluded. */
function viewerSources(dir: string, out: string[] = []): string[] {
	for (const entry of readdirSync(dir, { withFileTypes: true })) {
		const file = path.join(dir, entry.name);
		if (entry.isDirectory()) {
			viewerSources(file, out);
		} else if (/\.(ts|html)$/u.test(entry.name) && !/\.test(-support)?\.ts$/u.test(entry.name)) {
			out.push(file);
		}
	}
	return out;
}

/** Transition layers drawn inside the show stage, which addresses them by id. */
const TRANSITION_LAYERS = [
	'presentation-transition-overlay.component.ts',
	'morph-extra-layers.component.ts',
	'fragmented-transition-layer.component.ts',
];

describe('slide canvases take their element ids from the shared rule', () => {
	it('no call site sets the override', () => {
		const offenders = viewerSources(VIEWER_DIR).filter((file) =>
			readFileSync(file, 'utf8').includes('[exposeElementIds]'),
		);
		expect(offenders.map((file) => path.relative(VIEWER_DIR, file))).toStrictEqual([]);
	});

	it.each(TRANSITION_LAYERS)('%s keeps the ids on every layer', (file) => {
		const source = componentSource(VIEWER_DIR, file);
		const canvases = source.match(/<pptx-slide-canvas\b/gu) ?? [];
		const preserved = source.match(/\[preserveElementIds\]="true"/gu) ?? [];
		expect(canvases.length).toBeGreaterThan(0);
		expect(preserved).toHaveLength(canvases.length);
	});

	it('marks the live presentation stage as presenting', () => {
		const overlay = componentSource(VIEWER_DIR, 'presentation-overlay.component.ts');
		expect(overlay).toContain('[interactive]="false"');
		expect(overlay).toContain('[presenting]="true"');
	});

	it('turns the per-element gate off in the layout gallery, which draws elements directly', () => {
		expect(componentSource(VIEWER_DIR, 'ribbon-layout-gallery.component.ts')).toContain(
			'[exposeElementId]="false"',
		);
		expect(componentSource(VIEWER_DIR, 'element-renderer.component.ts')).toContain(
			'readonly exposeElementId = input<boolean>(true);',
		);
	});
});
