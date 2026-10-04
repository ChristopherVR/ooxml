import { fileURLToPath } from 'node:url';

// Tests run against the sibling ooxml checkout's source, like the demos (no build needed).
const ooxml = fileURLToPath(new URL(process.env.OOXML_DIR ?? '../ooxml-core/', import.meta.url)).replaceAll(
	String.fromCharCode(92),
	'/',
);

export default {
	resolve: {
		dedupe: ['yjs', 'lib0', 'y-protocols', 'lit'],
		alias: [
			{ find: /^ooxml-core\/(.+)$/, replacement: `${ooxml}src/$1/index.ts` },
			{ find: /^ooxml-ui$/, replacement: `${ooxml}packages/ui/src/index.ts` },
			{ find: /^teams-viewer$/, replacement: fileURLToPath(new URL('./packages/web-component/src/index.ts', import.meta.url)) },
		],
	},
	server: { fs: { strict: false } },
	test: {
		globals: true,
		environment: 'jsdom',
		include: ['packages/*/src/**/*.test.ts'],
		exclude: ['node_modules', 'dist'],
		testTimeout: 15000,
	},
};
