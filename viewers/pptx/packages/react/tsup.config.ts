import { defineConfig } from 'tsup';
import type { Options } from 'tsup';

import pkg from './package.json' with { type: 'json' };

/**
 * `pptx-viewer-core` is a thin `export * from 'ooxml-core/pptx'`. Bundling it turns that star
 * re-export into a runtime namespace copy (`__reExport`), which makes every export of
 * `ooxml-core/pptx` reachable for a host bundler, so it can no longer tree-shake the core.
 * Rewriting the specifier to the external `ooxml-core/pptx` keeps the imports static and named.
 */
const pptxViewerCoreAsOoxmlCore: NonNullable<Options['esbuildPlugins']>[number] = {
	name: 'pptx-viewer-core-as-ooxml-core',
	setup(build) {
		build.onResolve({ filter: /^pptx-viewer-core$/ }, () => ({
			path: 'ooxml-core/pptx',
			external: true,
		}));
	},
};

export default defineConfig((options) => ({
	entry: ['src/index.ts', 'src/viewer/index.ts', 'src/i18n.ts', 'src/internals.ts'],
	format: ['esm', 'cjs'],
	minify: false,
	// File > Account's About panel reads this to show the installed package
	// version without bundling `package.json` itself.
	define: {
		__PPTX_PACKAGE_VERSION__: JSON.stringify(pkg.version),
	},
	// Inline the .d.ts of the bundled internal workspace packages so the
	// published types resolve standalone: consumers don't need (and for
	// `pptx-viewer-shared`, can't get) those packages from npm. Mirrors the
	// runtime `noExternal` below and the Vue package's dts `bundledPackages`.
	dts: false,
	splitting: true,
	sourcemap: false,
	clean: !options.watch,
	// `ooxml-ui` depends on `ooxml-core`, so inlining the core here as well would ship it twice.
	external: [
		/^ooxml-ui(?:\/|$)/,
		/^ooxml-core(?:\/|$)/,
		'react',
		'react-dom',
		'framer-motion',
		'lucide-react',
		'react-icons',
		'html2canvas-pro',
		'jspdf',
		'jszip',
		'fast-xml-parser',
		'clsx',
		'tailwind-merge',
		'i18next',
		'react-i18next',
		// Shared types inserted into a host-owned Y.Doc must use its Yjs runtime.
		'yjs',
		// Optional AI SDK peers, reachable only through the lazily-loaded AI
		// chat panel / `pptx-viewer-shared/ai` subpath. Kept external so the
		// dynamic `import('ai')` inside shared stays a real optional runtime
		// import and neither SDK is ever inlined into this bundle.
		'ai',
		'@ai-sdk/react',
		// Node-only canvas backend the core `.ppt` writer loads to rasterise an
		// SVG outside a browser; optional peer, never part of a browser bundle.
		'@napi-rs/canvas',
	],
	// The internal workspace package `pptx-viewer-core` (a thin re-export of
	// `ooxml-core/pptx`) is resolved to `ooxml-core/pptx` itself, so consumers
	// can install just `pptx-react-viewer`. The core, and with it `emf-converter`
	// and `mtx-decompressor`, stays external above.
	esbuildPlugins: [pptxViewerCoreAsOoxmlCore],
	treeshake: true,
	platform: 'browser',
}));
