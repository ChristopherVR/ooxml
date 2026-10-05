import { defineConfig } from 'tsdown';

export default defineConfig({
	entry: {
		'pptx/automation/index': 'src/core/pptx/automation/index.ts',
		'pptx/ui/index': 'src/core/pptx/ui/index.ts',
		'pptx/index': 'src/core/pptx/index.ts',
		'pptx/converter/index': 'src/core/pptx/converter/index.ts',
		'pptx/cli/index': 'src/core/pptx/cli/index.ts',
		'pptx/signature-node/index': 'src/core/pptx/signature-node/index.ts',
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
