import { defineConfig } from 'tsdown';

export default defineConfig({
	entry: {
		'pptx/automation/index': 'pptx/automation/index.ts',
		'pptx/ui/index': 'pptx/ui/index.ts',
		'pptx/index': 'pptx/index.ts',
		'pptx/converter/index': 'pptx/converter/index.ts',
		'pptx/cli/index': 'pptx/cli/index.ts',
		'pptx/signature-node/index': 'pptx/signature-node/index.ts',
		'pptx/smartart-layouts/index': 'pptx/smartart-layouts/index.ts',
	},
	format: ['esm'],
	outDir: '.types-pptx',
	tsconfig: 'tsconfig.pptx.json',
	dts: { emitDtsOnly: true },
	sourcemap: false,
	clean: true,
	treeshake: true,
	platform: 'neutral',
	outputOptions: { chunkFileNames: 'pptx/chunks/[name]-[hash].js' },
	deps: { alwaysBundle: [/^@christophervr\/ole2(?:\/|$)/] },
});
