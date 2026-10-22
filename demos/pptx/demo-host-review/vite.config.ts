import { resolve } from 'node:path';

import { defineConfig } from 'vite';

/**
 * The review hub's page. `server/index.ts` mounts this config as middleware on the hub's own HTTP
 * server, so the page, the API and the collaboration socket share one origin. As in the other
 * demos, the viewer packages are aliased to their source so edits hot-reload.
 */
const pkg = (...p: string[]) =>
	resolve(import.meta.dirname, '..', '..', '..', 'viewers', 'pptx', 'packages', ...p);

export default defineConfig({
	root: import.meta.dirname,
	publicDir: false,
	build: { chunkSizeWarningLimit: 2500 },
	resolve: {
		// Order matters: more specific subpath aliases must precede the bare ones.
		alias: [
			{ find: 'pptx-vanilla-viewer', replacement: pkg('vanilla', 'src', 'index.ts') },
			{
				find: 'pptx-viewer-core/converter',
				replacement: pkg('core', 'src', 'converter', 'index.ts'),
			},
			{ find: 'pptx-viewer-core/chart', replacement: pkg('core', 'src', 'chart', 'index.ts') },
			{ find: 'pptx-viewer-core/text', replacement: pkg('core', 'src', 'text', 'index.ts') },
			{
				find: 'pptx-viewer-core/geometry',
				replacement: pkg('core', 'src', 'geometry', 'index.ts'),
			},
			{ find: 'pptx-viewer-core/color', replacement: pkg('core', 'src', 'color', 'index.ts') },
			{ find: 'pptx-viewer-core/ui', replacement: pkg('core', 'src', 'ui', 'index.ts') },
			{
				find: 'pptx-viewer-core/smartart-layouts',
				replacement: pkg('core', 'src', 'smartart-layouts', 'index.ts'),
			},
			{ find: 'pptx-viewer-core/math', replacement: pkg('core', 'src', 'math', 'index.ts') },
			{ find: 'pptx-viewer-core', replacement: pkg('core', 'src', 'index.ts') },
			{ find: 'pptx-viewer-locales', replacement: pkg('locales', 'src', 'index.ts') },
			{ find: 'pptx-viewer-shared/i18n', replacement: pkg('shared', 'src', 'i18n', 'index.ts') },
			{ find: 'pptx-viewer-shared/ai', replacement: pkg('shared', 'src', 'ai', 'index.ts') },
			{ find: 'pptx-viewer-shared', replacement: pkg('shared', 'src', 'index.ts') },
		],
	},
	optimizeDeps: {
		include: ['jszip', 'fast-xml-parser'],
		// See demo-react/vite.config.ts: pre-bundling breaks emf-converter's optional canvas import.
		exclude: ['emf-converter'],
	},
});
