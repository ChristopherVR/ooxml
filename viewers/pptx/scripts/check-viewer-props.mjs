/**
 * Fails when a PowerPoint binding's public props, inputs, options or events drift from the
 * shared manifest (`src/ui/src/pptx/render/viewer-props-manifest.ts`).
 *
 *   bun scripts/check-viewer-props.mjs        (or: bun run check:viewer-props)
 *
 * Runs under Bun because the manifest is TypeScript. It replaces the advisory path heuristics of
 * `parity-check.mjs` for the question "does every binding expose the same viewer surface?": that
 * one cannot tell a deliberate difference from an accident, this one reads the real types and
 * accepts only the exceptions the manifest records.
 */
import { pathToFileURL } from 'node:url';

import { diffViewerSurface } from '../../../src/ui/src/pptx/render/viewer-props-manifest.ts';
import { extractBindings } from './viewer-props-extract.mjs';

/** Every drift between the bindings' current source and the manifest (empty when in step). */
export function findViewerPropDrift() {
	return diffViewerSurface(extractBindings());
}

export function reportViewerPropDrift(drift) {
	return [
		`Viewer props drift: ${drift.length} difference${drift.length === 1 ? '' : 's'} between the bindings and the manifest.`,
		...drift.map((entry) => `  - [${entry.binding}] ${entry.message}`),
		'',
		'Fix the binding, or update src/ui/src/pptx/render/viewer-props-manifest.ts (see AGENTS.md, Rule 1).',
	].join('\n');
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
	const drift = findViewerPropDrift();
	if (drift.length > 0) {
		console.error(reportViewerPropDrift(drift));
		process.exit(1);
	}
	console.log('Viewer props: all five bindings match the manifest.');
}
