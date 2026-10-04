import { fileURLToPath } from 'node:url';
import type { UserConfig } from 'vite';

// Local development against a sibling checkout of the ooxml repository. By default a demo
// resolves ooxml-core and ooxml-ui from their SOURCE (instant reload when you edit either);
// set TEAMS_USE_DIST=1 to use the built packages installed in node_modules instead.
const ooxml = fileURLToPath(new URL(process.env.OOXML_DIR ?? '../../ooxml-core/', import.meta.url));
const useSource = process.env.TEAMS_USE_DIST !== '1';

export function sharedConfig(port: number): UserConfig {
	return {
		resolve: {
			// Yjs and Lit break (instanceof, double registration) if two copies load: force one.
			dedupe: ['yjs', 'lib0', 'y-protocols', 'lit', 'react', 'react-dom'],
			alias: useSource
				? [
						{ find: /^ooxml-core\/(.+)$/, replacement: `${ooxml}src/$1/index.ts` },
						{ find: /^ooxml-ui$/, replacement: `${ooxml}packages/ui/src/index.ts` },
					]
				: [],
		},
		server: { port, strictPort: true, host: '127.0.0.1', fs: { allow: ['..', ooxml] } },
	};
}
