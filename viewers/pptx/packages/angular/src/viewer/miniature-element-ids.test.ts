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

	it('every element renderer mounted outside a slide canvas turns the gate off', () => {
		const offenders: string[] = [];
		for (const file of viewerSources(VIEWER_DIR)) {
			const relative = path.relative(VIEWER_DIR, file);
			if (
				relative.startsWith('slide-canvas.component') ||
				relative.startsWith('element-renderer')
			) {
				continue;
			}
			for (const [mount] of readFileSync(file, 'utf8').matchAll(
				/<pptx-element-renderer\b[^>]*>/gu,
			)) {
				if (!mount.includes('[exposeElementId]="false"')) {
					offenders.push(relative);
				}
			}
		}
		expect(offenders).toStrictEqual([]);
	});
});

/**
 * The canvas hands its decision to `ElementRendererComponent.exposeElementId`,
 * which threads it into every per-type renderer. A renderer that binds
 * `data-element-id` straight from the element (zoom and 3D SmartArt once did)
 * ignores it, so every element id binding has to go through `elementIdAttr()`.
 */
describe('every element renderer honours the per-element gate', () => {
	/** Remote collaborators' selection boxes: an editing-canvas overlay, keyed per box. */
	const NOT_ELEMENTS = new Set(['remote-selection-overlay.component.ts']);

	it('binds data-element-id only through elementIdAttr()', () => {
		const offenders: string[] = [];
		for (const file of viewerSources(VIEWER_DIR)) {
			const relative = path.relative(VIEWER_DIR, file);
			if (NOT_ELEMENTS.has(relative)) {
				continue;
			}
			for (const [binding] of readFileSync(file, 'utf8').matchAll(
				/\[attr\.data-element-id\]="[^"]*"/gu,
			)) {
				if (binding !== '[attr.data-element-id]="elementIdAttr()"') {
					offenders.push(`${relative}: ${binding}`);
				}
			}
		}
		expect(offenders).toStrictEqual([]);
	});

	it('passes the gate to every per-type renderer the graphics dispatcher mounts', () => {
		const source = readFileSync(
			path.join(VIEWER_DIR, 'element-renderer-graphics.component.html'),
			'utf8',
		);
		const gated: string[] = [];
		const wrapped: string[] = [];
		for (const [mount, name] of source.matchAll(/<pptx-([a-z0-9-]+-renderer)\b[^>]*>/gu)) {
			(mount.includes('[exposeElementId]="exposeElementId()"') ? gated : wrapped).push(name!);
		}
		// These position (and mark) their own root, so they take the gate.
		expect(gated.sort()).toStrictEqual([
			'content-part-renderer',
			'ink-renderer',
			'media-renderer',
			'model3d-renderer',
			'smart-art-3d-renderer',
			'zoom-renderer',
		]);
		// These draw inside a box the dispatcher marks itself, with `elementIdAttr()`.
		expect(wrapped.sort()).toStrictEqual(['ole-renderer', 'smart-art-renderer', 'table-renderer']);
	});
});
