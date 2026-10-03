import { defineConfig } from 'tsdown';

export default defineConfig({
	entry: {
		'pptx/automation/index': 'src/pptx/automation/index.ts',
		'pptx/index': 'src/pptx/index.ts',
		'pptx/converter/index': 'src/pptx/converter/index.ts',
		'pptx/cli/index': 'src/pptx/cli/index.ts',
		'pptx/signature-node/index': 'src/pptx/signature-node/index.ts',
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
