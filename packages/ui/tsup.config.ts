import { defineConfig } from 'tsup';

// ESM only: custom elements need a browser (or a DOM) anyway. Declarations come from tsc
// (tsconfig.build.json). The core package stays external: it is a real dependency.
export default defineConfig({
	entry: {
		index: 'src/index.ts',
		theme: 'src/theme.ts',
		icons: 'src/icons.ts',
		controls: 'src/controls.ts',
		presence: 'src/presence.ts',
		smartart: 'src/smartart.ts',
	},
	outDir: 'dist',
	tsconfig: 'tsconfig.build.json',
	format: ['esm'],
	dts: false,
	// Shared helpers (registry, styles, icons) would otherwise be duplicated per entry, which
	// would break the idempotent registration across entries.
	splitting: true,
	clean: false,
	treeshake: true,
	platform: 'browser',
	external: [/^@christophervr\/ooxml-core/],
});
