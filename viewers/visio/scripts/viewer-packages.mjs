// The seven publishable packages of this viewer, as the build and package tests use them. The
// release table lives in the repository's scripts/viewer-packages.mjs (the root planner releases
// them); this view only strips the `viewers/visio/` prefix and the mcp server, which ships source.
import { VIEWER_PACKAGES as ALL } from '../../../scripts/viewer-packages.mjs';

export const VIEWER_PACKAGES = Object.fromEntries(
	Object.entries(ALL)
		.filter(([key]) => key.startsWith('visio-') && key !== 'visio-mcp')
		.map(([key, meta]) => [
			key.slice('visio-'.length),
			{ dir: meta.dir.replace(/^viewers\/visio\//u, ''), npm: meta.npm },
		]),
);
