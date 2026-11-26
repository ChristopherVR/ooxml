import { resolve } from 'node:path';

import { defineConfig } from 'vitest/config';

import { testWorkers } from '../../../../scripts/test-workers.mjs';
import { isolationProjects } from '../../scripts/vitest-isolation.mjs';

// TestBed files share workers and native dependencies. The fresh-module project
// re-evaluates binding components because fixtures alter JIT signal-input metadata.
export default defineConfig({
	server: { fs: { allow: [resolve(import.meta.dirname, '../../../..')] } },
	resolve: {
		// Keep leaf-module mocks attached to the renderer implementation rather
		// than a separately bundled copy of the same helper.
		alias: {
			'ooxml-ui/pptx': resolve(import.meta.dirname, '../../../../src/ui/src/pptx'),
			// The `/ai` secondary entry point imports the package root by name.
			'pptx-angular-viewer': resolve(import.meta.dirname, 'src/public-api.ts'),
			// Test against the locale sources, not a possibly stale dist.
			'pptx-viewer-locales': resolve(import.meta.dirname, '../locales/src/index.ts'),
		},
	},
	test: {
		environment: 'happy-dom',
		maxWorkers: testWorkers(),
		projects: isolationProjects('angular', import.meta.dirname, 'happy-dom', [], {
			testBedSetup: './vitest-test-bed.setup.ts',
		}),
		globals: true,
		setupFiles: ['./vitest-setup.ts'],
	},
});
