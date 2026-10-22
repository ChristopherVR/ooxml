// The "business system": sign-in, review items, permissions, versions and the AI route, with the
// collaboration relay on the same origin. Everything the editor needs from the host goes through
// the four calls the client makes: the review record, the source deck, a room ticket and (for the
// owner) the write-back of saved versions.
import { createServer } from 'node:http';
import type { IncomingMessage, ServerResponse } from 'node:http';
import { Readable } from 'node:stream';

import { aiSettings, streamReviewChat } from './ai-route';
import type { AiSettings } from './ai-route';
import { ACCOUNTS, canWrite, createDirectory, viewerRole } from './directory';
import type { Permission } from './directory';
import { createRelay } from './relay';
import { SIGN_IN_TTL_MS, TICKET_TTL_MS, signToken, verifySignIn, verifyTicket } from './tickets';

const COOKIE = 'review_hub';
const MAX_UPLOAD_BYTES = 64 * 1024 * 1024;
const PERMISSIONS = new Set<Permission>(['review', 'view', 'none']);

class HttpError extends Error {
	constructor(
		readonly status: number,
		message: string,
	) {
		super(message);
	}
}

async function readBody(request: IncomingMessage, limit: number): Promise<Buffer> {
	const chunks: Buffer[] = [];
	let size = 0;
	for await (const chunk of request as AsyncIterable<Buffer>) {
		size += chunk.length;
		if (size > limit) throw new HttpError(413, 'Too large');
		chunks.push(chunk);
	}
	return Buffer.concat(chunks);
}

const json = (response: ServerResponse, status: number, value: unknown): void => {
	response.writeHead(status, { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' });
	response.end(JSON.stringify(value));
};

const bytes = (response: ServerResponse, data: Uint8Array, fileName: string): void => {
	response.writeHead(200, {
		'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
		'Content-Disposition': `attachment; filename="${fileName}"`,
		'Cache-Control': 'no-store',
	});
	response.end(data);
};

async function sendWeb(response: ServerResponse, web: Response): Promise<void> {
	response.writeHead(web.status, Object.fromEntries(web.headers));
	if (!web.body) return void response.end();
	Readable.fromWeb(web.body as import('node:stream/web').ReadableStream).pipe(response);
}

function cookie(request: IncomingMessage, name: string): string {
	for (const part of (request.headers.cookie ?? '').split(';')) {
		const [key, ...value] = part.trim().split('=');
		if (key === name) return decodeURIComponent(value.join('='));
	}
	return '';
}

export function createReviewHub(options: {
	secret: string;
	source: Uint8Array;
	ai?: AiSettings | null;
	/** Requests the hub does not route (the Vite dev middleware in development). */
	fallback?: (request: IncomingMessage, response: ServerResponse) => void;
}) {
	const { secret } = options;
	const ai = options.ai === undefined ? aiSettings() : options.ai;
	const directory = createDirectory(options.source);
	const relay = createRelay({ verify: (token) => verifyTicket(secret, token) });

	const signedIn = (request: IncomingMessage): string => {
		const session = verifySignIn(secret, cookie(request, COOKIE));
		if (!session || !directory.account(session.account)) throw new HttpError(401, 'Sign in first');
		return session.account;
	};

	/** The review and the caller's permission; 404 hides items the caller may not see. */
	const access = (request: IncomingMessage, reviewId: string) => {
		const account = signedIn(request);
		const review = directory.review(reviewId);
		const permission = directory.permission(reviewId, account);
		if (!review || permission === 'none') throw new HttpError(404, 'No such review');
		return { account, review, permission };
	};

	async function route(
		request: IncomingMessage,
		response: ServerResponse,
		url: URL,
	): Promise<void> {
		const method = request.method ?? 'GET';
		// Every state-changing call carries this header, which a cross-site form cannot set.
		if (method !== 'GET' && request.headers['x-review-hub'] !== '1') {
			throw new HttpError(403, 'Missing request header');
		}
		const path = url.pathname;
		if (path === '/api/accounts' && method === 'GET') {
			return json(
				response,
				200,
				ACCOUNTS.map(({ id, name, title, color }) => ({ id, name, title, color })),
			);
		}
		if (path === '/api/session' && method === 'POST') {
			// Demo sign-in: pick an account. Replace with your identity provider.
			const body = JSON.parse((await readBody(request, 4096)).toString() || '{}') as {
				account?: string;
			};
			if (!body.account || !directory.account(body.account))
				throw new HttpError(400, 'Unknown account');
			const token = signToken(secret, { account: body.account, exp: Date.now() + SIGN_IN_TTL_MS });
			response.setHeader(
				'Set-Cookie',
				`${COOKIE}=${token}; HttpOnly; SameSite=Strict; Path=/; Max-Age=${SIGN_IN_TTL_MS / 1000}`,
			);
			return json(response, 200, { ok: true });
		}
		if (path === '/api/session' && method === 'DELETE') {
			response.setHeader('Set-Cookie', `${COOKIE}=; HttpOnly; SameSite=Strict; Path=/; Max-Age=0`);
			return json(response, 200, { ok: true });
		}
		if (path === '/api/me' && method === 'GET') {
			const account = directory.account(signedIn(request));
			return json(response, 200, account);
		}
		if (path === '/api/ai/chat' && method === 'POST') {
			const reviewId = url.searchParams.get('review') ?? '';
			const { account, review } = access(request, reviewId);
			if (!ai || !directory.canUseAi(reviewId, account)) throw new HttpError(403, 'AI not allowed');
			const body: unknown = JSON.parse((await readBody(request, 4 * 1024 * 1024)).toString());
			directory.log(review, account, 'asked the assistant');
			const reviewer = directory.account(account)?.name ?? account;
			return sendWeb(
				response,
				await streamReviewChat(ai, body, { reviewTitle: review.title, reviewer }),
			);
		}

		const match =
			/^\/api\/reviews\/([\w-]+)(?:\/(source|ticket|versions|access)(?:\/([\w-]+))?)?$/u.exec(path);
		if (!match) throw new HttpError(404, 'Not found');
		const [, reviewId = '', part, item] = match;
		const { account, review, permission } = access(request, reviewId);

		if (!part && method === 'GET') {
			const here = new Set(relay.online(reviewId));
			return json(response, 200, {
				id: review.id,
				title: review.title,
				fileName: review.fileName,
				permission,
				ai: Boolean(ai) && directory.canUseAi(reviewId, account),
				people: ACCOUNTS.map(({ id, name, title, color }) => ({
					id,
					name,
					title,
					color,
					permission: directory.permission(reviewId, id),
					online: here.has(id),
				})),
				versions: review.versions.map(({ number, savedAt, savedBy }) => ({
					number,
					savedAt,
					savedBy,
				})),
				activity: review.activity,
			});
		}
		if (part === 'source' && method === 'GET')
			return bytes(response, review.source, review.fileName);
		if (part === 'ticket' && method === 'POST') {
			const exp = Date.now() + TICKET_TTL_MS;
			const user = directory.account(account);
			return json(response, 200, {
				ticket: signToken(secret, { account, review: reviewId, permission, exp }),
				expiresAt: exp,
				permission,
				role: viewerRole(permission),
				// Only one writer seeds an empty room; everybody else joins and waits for it. A refresh
				// of a mounted session's ticket never claims the seed.
				intent:
					canWrite(permission) &&
					!url.searchParams.has('refresh') &&
					relay.claimSeed(reviewId, account)
						? 'create'
						: 'join',
				user: { name: user?.name, color: user?.color },
			});
		}
		if (part === 'versions' && method === 'POST' && !item) {
			const version = directory.addVersion(
				reviewId,
				account,
				await readBody(request, MAX_UPLOAD_BYTES),
			);
			if (!version) throw new HttpError(403, 'Only the owner saves versions');
			return json(response, 200, { number: version.number });
		}
		if (part === 'versions' && method === 'GET' && item) {
			const version = review.versions.find((entry) => String(entry.number) === item);
			if (!version) throw new HttpError(404, 'No such version');
			return bytes(
				response,
				version.bytes,
				review.fileName.replace(/\.pptx$/u, `-v${version.number}.pptx`),
			);
		}
		if (part === 'access' && method === 'PUT' && item) {
			const body = JSON.parse((await readBody(request, 4096)).toString() || '{}') as {
				permission?: Permission;
			};
			const next = body.permission;
			if (
				!next ||
				!PERMISSIONS.has(next) ||
				!directory.setPermission(reviewId, account, item, next)
			) {
				throw new HttpError(403, 'Cannot change that access');
			}
			// Open sockets were admitted with the old access: close them so they come back with a new ticket.
			relay.revoke(reviewId, item);
			return json(response, 200, { ok: true });
		}
		throw new HttpError(405, 'Method not allowed');
	}

	const server = createServer((request, response) => {
		const url = new URL(request.url ?? '/', 'http://hub');
		if (!url.pathname.startsWith('/api/')) {
			if (options.fallback) return options.fallback(request, response);
			response.writeHead(404).end();
			return;
		}
		route(request, response, url).catch((error: unknown) => {
			if (response.headersSent) return void response.destroy();
			const status = error instanceof HttpError ? error.status : 500;
			json(response, status, { error: error instanceof Error ? error.message : 'Server error' });
		});
	});
	server.on('upgrade', (request, socket, head) => {
		// Other upgrades (the Vite HMR socket in development) are handled by their own listeners.
		relay.upgrade(request, socket, head);
	});

	return {
		server,
		directory,
		async close(): Promise<void> {
			await relay.close();
			server.close();
		},
	};
}
