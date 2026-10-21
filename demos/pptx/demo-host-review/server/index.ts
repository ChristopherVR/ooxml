// Starts the review hub and the Vite dev server for its page on one origin, so the sign-in cookie,
// the API and the collaboration socket share it: `bun run start`, then open the printed URL.
import { randomBytes } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { createServer as createVite } from 'vite';

import { createReviewHub } from './app';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const port = Number(process.env.PORT) || 4190;
const source = readFileSync(resolve(root, '../../../e2e/pptx/fixtures/sample-deck.pptx'));

const hub = createReviewHub({
	// Tickets and sign-ins are signed with this. Set one to keep sessions across restarts.
	secret: process.env.REVIEW_HUB_SECRET || randomBytes(32).toString('hex'),
	source: new Uint8Array(source),
	fallback: (request, response) => vite.middlewares(request, response),
});
const vite = await createVite({
	root,
	configFile: resolve(root, 'vite.config.ts'),
	server: { middlewareMode: true, hmr: { server: hub.server } },
	appType: 'spa',
});

hub.server.listen(port, '127.0.0.1', () => {
	console.log(`Review hub: http://localhost:${port}/`);
	console.log('Open it in two browsers (or a private window) and sign in as different people.');
});

const stop = (): void => {
	void Promise.all([hub.close(), vite.close()]).finally(() => process.exit(0));
};
process.on('SIGINT', stop);
process.on('SIGTERM', stop);
