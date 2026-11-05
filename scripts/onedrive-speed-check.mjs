/**
 * Download speed measurement for the OneDrive integration. Live mode needs a short-lived Graph
 * token in the environment; --mock measures a local imitation. Tokens, sharing links, drive paths
 * and temporary download URLs are never printed or stored (see docs/onedrive-testing.md).
 */
import {
	SPEED_USAGE,
	encodeDrivePath,
	entryId,
	fmtBytes,
	fmtMs,
	fmtRate,
	interpret,
	notFoundHint,
	parseArgs,
	renderTable,
	stats,
	throughput,
} from './onedrive-speed-lib.mjs';
import { checkToken, normalizeToken } from './onedrive-token-check.mjs';

const GRAPH = 'https://graph.microsoft.com/v1.0';
// Same field list as the library's open request, so the metadata payload is comparable.
const OPEN_SELECT =
	'id,name,size,eTag,webUrl,folder,file,package,remoteItem,parentReference,@microsoft.graph.downloadUrl';
const LIST_SELECT = 'id,name,size,folder,file,lastModifiedDateTime,parentReference';
const now = () => performance.now();

class SafeError extends Error {}

async function main() {
	const options = parseArgs(process.argv.slice(2));
	if (options.help) return console.log(SPEED_USAGE);
	const lib = await import('../src/core/dist/embedding/onedrive/index.mjs');
	let mock;
	let baseFetch = globalThis.fetch.bind(globalThis);
	let token;
	let shareUrl;
	let drivePath;
	let folderPath;
	if (options.mock) {
		const { startMockGraph } = await import('./onedrive-speed-mock.mjs');
		mock = await startMockGraph(options);
		baseFetch = mock.fetch;
		token = 'mock-token';
		shareUrl = mock.shareUrl;
		drivePath = mock.path;
		folderPath = mock.folder;
	} else {
		token = normalizeToken(process.env.OOXML_ONEDRIVE_ACCESS_TOKEN);
		shareUrl = process.env.OOXML_ONEDRIVE_SHARE_URL || undefined;
		drivePath = process.env.OOXML_ONEDRIVE_PATH || undefined;
		folderPath = process.env.OOXML_ONEDRIVE_FOLDER || undefined;
		if (token && options.list) {
			try {
				await checkToken(token);
			} catch (error) {
				throw new SafeError(error.message);
			}
			return await listNames({ token, baseFetch, folderPath });
		}
		if (!token || (!shareUrl && !drivePath))
			throw new SafeError(
				'Set OOXML_ONEDRIVE_ACCESS_TOKEN and OOXML_ONEDRIVE_SHARE_URL and/or OOXML_ONEDRIVE_PATH (see --help).',
			);
		try {
			await checkToken(token);
		} catch (error) {
			throw new SafeError(error.message);
		}
	}
	try {
		const targets = [];
		if (shareUrl) {
			const enc = lib.encodeOneDriveSharingUrl(shareUrl);
			targets.push({
				label: 'sharing link',
				meta: `/shares/${enc}/driveItem?$select=${OPEN_SELECT}`,
				content: `/shares/${enc}/driveItem/content`,
				headers: { Prefer: 'redeemSharingLink' },
				openLibrary: (client, input) => client.open(shareUrl, input),
			});
		}
		if (drivePath) {
			const enc = encodeDrivePath(drivePath);
			targets.push({
				label: 'drive path',
				meta: `/me/drive/root:/${enc}?$select=${OPEN_SELECT}`,
				content: `/me/drive/root:/${enc}:/content`,
				headers: {},
				openLibrary: (client, input, ref) => client.openItem(ref, input),
			});
		}
		const ctx = { options, lib, token, baseFetch, folderPath };
		const results = [];
		for (const target of targets) results.push(await measureTarget(ctx, target));
		const listing = await measureListing(ctx);
		report(options, results, listing);
	} finally {
		await mock?.close();
	}
}

/** --list: the names in the drive root (or the folder), to find the exact path to measure. */
async function listNames(ctx) {
	const select = '$select=name,folder,file,size&$top=200';
	const path = ctx.folderPath
		? `/me/drive/root:/${encodeDrivePath(ctx.folderPath)}:/children?${select}`
		: `/me/drive/root/children?${select}`;
	const { data } = await rawGraph(ctx, path, { what: ctx.folderPath ? 'folder' : undefined });
	const items = Array.isArray(data?.value) ? data.value : [];
	console.log(ctx.folderPath ? 'Items in OOXML_ONEDRIVE_FOLDER:' : 'Items in your OneDrive root:');
	console.log(
		renderTable([
			['name', 'type', 'size'],
			...items.map((item) => [
				String(item.name ?? ''),
				item.folder ? 'folder' : 'file',
				item.folder ? `${item.folder.childCount ?? 0} items` : fmtBytes(Number(item.size ?? 0)),
			]),
		]),
	);
	console.log(
		'Use a file as OOXML_ONEDRIVE_PATH (folder/file.pptx), or a folder as OOXML_ONEDRIVE_FOLDER with --list to look inside it.',
	);
}

/** Raw Graph call: time to response headers, total time and body size. */
async function rawGraph(ctx, path, { headers = {}, redirect = 'error', what } = {}) {
	const started = now();
	const response = await ctx.baseFetch(GRAPH + path, {
		headers: { Authorization: `Bearer ${ctx.token}`, ...headers },
		redirect,
		credentials: 'omit',
		referrerPolicy: 'no-referrer',
		signal: AbortSignal.timeout(120000),
	});
	const headersMs = now() - started;
	if (redirect === 'manual' && response.status >= 300 && response.status < 400) {
		await response.body?.cancel();
		const location = response.headers.get('Location');
		if (!location) throw new SafeError('Graph redirected without a download location.');
		return { location, headersMs, totalMs: now() - started };
	}
	if (!response.ok) {
		let code = '';
		try {
			code = String((await response.json())?.error?.code ?? '');
		} catch {
			// No JSON body.
		}
		const lines = [
			`${what ? `Looking up the ${what}` : 'A Graph request'} failed with HTTP ${response.status}${code ? ` (${code})` : ''}.`,
		];
		const hint = notFoundHint(what, response.status);
		if (hint) lines.push(hint);
		throw new SafeError(lines.join(String.fromCharCode(10)));
	}
	const body = new Uint8Array(await response.arrayBuffer());
	const totalMs = now() - started;
	return {
		data: JSON.parse(new TextDecoder().decode(body)),
		headersMs,
		totalMs,
		bytes: body.length,
	};
}

/** Raw download of a pre-authenticated URL (never printed): first byte, total, chunks. */
async function rawDownload(ctx, url) {
	const started = now();
	const response = await ctx.baseFetch(url, {
		credentials: 'omit',
		referrerPolicy: 'no-referrer',
		signal: AbortSignal.timeout(300000),
	});
	const headersMs = now() - started;
	if (!response.ok) throw new SafeError(`The download host answered HTTP ${response.status}.`);
	const reader = response.body.getReader();
	let bytes = 0;
	let chunks = 0;
	let firstByteMs = NaN;
	for (;;) {
		const { value, done } = await reader.read();
		if (done) break;
		if (!chunks) firstByteMs = now() - started;
		chunks++;
		bytes += value.byteLength;
	}
	const totalMs = now() - started;
	return { headersMs, firstByteMs, totalMs, bytes, chunks, mbps: throughput(bytes, totalMs) };
}

/** Library fetch that only timestamps response headers per host, so it adds no body overhead. */
function tracedFetch(ctx, trace) {
	return async (input, init) => {
		const host = new URL(String(input)).hostname;
		const started = now();
		const response = await ctx.baseFetch(input, init);
		trace.push({ graph: host === 'graph.microsoft.com', headersMs: now() - started });
		return response;
	};
}

async function timed(run) {
	const started = now();
	try {
		await run();
		return { totalMs: now() - started };
	} catch (error) {
		// Library messages are written to be safe to show; anything else is reduced to a stub.
		return { totalMs: NaN, error: error?.name === 'OneDriveError' ? error.message : 'failed' };
	}
}

async function measureTarget(ctx, target) {
	const { options, lib } = ctx;
	const runs = [];
	let ref;
	const total = options.runs + (options.warmup ? 1 : 0);
	for (let i = 0; i < total; i++) {
		const run = {};
		// The token here is static: acquiring it is a function call, not a sign-in.
		const t0 = now();
		await Promise.resolve(ctx.token);
		run.tokenMs = now() - t0;
		const meta = await rawGraph(ctx, target.meta, { headers: target.headers, what: target.label });
		run.meta = { headersMs: meta.headersMs, totalMs: meta.totalMs, bytes: meta.bytes };
		let item = meta.data;
		let downloadUrl = item['@microsoft.graph.downloadUrl'];
		// Some responses (personal OneDrive links among them) drop the download URL when fields are
		// selected; ask for the full item, as the library does, and count it in the metadata time.
		if (typeof downloadUrl !== 'string' && item.id && item.parentReference?.driveId) {
			const full = await rawGraph(
				ctx,
				`/drives/${encodeURIComponent(item.parentReference.driveId)}/items/${encodeURIComponent(item.id)}`,
				{ what: target.label },
			);
			run.meta.totalMs += full.totalMs;
			run.meta.bytes += full.bytes;
			run.fullItem = true;
			item = full.data;
			downloadUrl = item['@microsoft.graph.downloadUrl'];
		}
		if (typeof downloadUrl !== 'string')
			throw new SafeError('Graph did not offer a download for this file, even for the full item.');
		run.download = await rawDownload(ctx, downloadUrl);
		run.rawMs = run.meta.totalMs + run.download.totalMs;
		ref = { driveId: item.parentReference?.driveId, itemId: item.id };
		const trace = [];
		const getAccessToken = async () => ctx.token;
		const client = lib.createOneDriveClient({
			getAccessToken,
			fetch: tracedFetch(ctx, trace),
			maxBytes: 2 * 1024 * 1024 * 1024,
			timeoutMs: 300000,
		});
		let progress = 0;
		let firstProgressMs = NaN;
		const libStart = now();
		const onProgress = () => {
			if (!progress++) firstProgressMs = now() - libStart;
		};
		const opened = await timed(() => target.openLibrary(client, { onProgress }, ref));
		run.library = {
			...opened,
			progressCalls: progress,
			firstProgressMs,
			graphHeadersMs: trace.filter((t) => t.graph).reduce((s, t) => s + t.headersMs, 0),
			downloadHeadersMs: trace.filter((t) => !t.graph).reduce((s, t) => s + t.headersMs, 0),
		};
		const providerTrace = [];
		const provider = lib.createOneDriveFileProvider({
			getAccessToken,
			fetch: tracedFetch(ctx, providerTrace),
			maxBytes: 2 * 1024 * 1024 * 1024,
			timeoutMs: 300000,
			account: { name: 'Speed check' },
			signIn: async () => ({ name: 'Speed check' }),
		});
		const entry = {
			id: entryId(ref),
			name: String(item.name ?? 'file'),
			type: 'file',
			kind: String(item.name ?? '')
				.split('.')
				.pop(),
		};
		run.provider = await timed(() => provider.open(entry, {}));
		if (options.content) {
			try {
				const redirect = await rawGraph(ctx, target.content, {
					headers: target.headers,
					redirect: 'manual',
				});
				const content = await rawDownload(ctx, redirect.location);
				run.content = {
					redirectMs: redirect.totalMs,
					downloadMs: content.totalMs,
					totalMs: redirect.totalMs + content.totalMs,
					bytes: content.bytes,
				};
			} catch (error) {
				run.content = {
					totalMs: NaN,
					error: error instanceof SafeError ? error.message : 'failed',
				};
			}
		}
		if (!options.warmup || i > 0) runs.push(run);
	}
	if (runs.some((run) => run.fullItem))
		console.error(
			`Note (${target.label}): Graph left the download URL out of the selected response, so the metadata time includes a second request for the full item. The library does the same.`,
		);
	return { label: target.label, runs };
}

async function measureListing(ctx) {
	const { lib, options } = ctx;
	const client = lib.createOneDriveClient({
		getAccessToken: async () => ctx.token,
		fetch: ctx.baseFetch,
	});
	const steps = [
		['list()', () => client.list(undefined)],
		['recent()', () => client.recent()],
	];
	if (ctx.folderPath) {
		const folder = await rawGraph(
			ctx,
			`/me/drive/root:/${encodeDrivePath(ctx.folderPath)}?$select=${LIST_SELECT}`,
			{ what: 'folder' },
		);
		const ref = { driveId: folder.data.parentReference?.driveId, itemId: folder.data.id };
		steps.push(['list(folder)', () => client.list(ref)]);
	}
	const out = [];
	for (const [name, run] of steps) {
		const times = [];
		let entries = 0;
		let error;
		for (let i = 0; i < options.runs; i++) {
			const started = now();
			try {
				entries = (await run()).entries.length;
				times.push(now() - started);
			} catch (e) {
				error = e?.name === 'OneDriveError' ? e.message : 'failed';
				break;
			}
		}
		out.push({ name, entries, times, stats: stats(times), ...(error ? { error } : {}) });
	}
	return out;
}

const col = (runs, pick) => runs.map(pick).filter((value) => Number.isFinite(value));

function summarize(target) {
	const r = target.runs;
	const s = (pick) => stats(col(r, pick));
	return {
		tokenMs: s((x) => x.tokenMs),
		metaHeadersMs: s((x) => x.meta.headersMs),
		metaTotalMs: s((x) => x.meta.totalMs),
		metaBytes: s((x) => x.meta.bytes),
		downloadHeadersMs: s((x) => x.download.headersMs),
		downloadFirstByteMs: s((x) => x.download.firstByteMs),
		downloadTotalMs: s((x) => x.download.totalMs),
		downloadBytes: s((x) => x.download.bytes),
		downloadMBps: s((x) => x.download.mbps),
		downloadChunks: s((x) => x.download.chunks),
		rawMs: s((x) => x.rawMs),
		libraryMs: s((x) => x.library.totalMs),
		libraryGraphHeadersMs: s((x) => x.library.graphHeadersMs),
		libraryDownloadHeadersMs: s((x) => x.library.downloadHeadersMs),
		libraryFirstProgressMs: s((x) => x.library.firstProgressMs),
		libraryProgressCalls: s((x) => x.library.progressCalls),
		providerMs: s((x) => x.provider.totalMs),
		contentRedirectMs: s((x) => x.content?.redirectMs ?? NaN),
		contentTotalMs: s((x) => x.content?.totalMs ?? NaN),
	};
}

function report(options, results, listing) {
	const summaries = results.map((target) => ({ target, summary: summarize(target) }));
	if (options.json) {
		console.log(
			JSON.stringify(
				{
					mode: options.mock
						? {
								mock: true,
								sizeMB: options.size,
								latencyMs: options.latency,
								bandwidthMBps: options.bandwidth,
							}
						: { mock: false },
					runs: options.runs,
					warmupDiscarded: options.warmup,
					targets: summaries.map(({ target, summary }) => ({
						label: target.label,
						runs: target.runs,
						summary,
					})),
					listing,
				},
				null,
				2,
			),
		);
		return;
	}
	const out = [];
	out.push(
		options.mock
			? `OneDrive speed check (MOCK: ${options.size} MB, ${options.latency} ms latency, ${options.bandwidth > 0 ? options.bandwidth + ' MB/s cap' : 'no bandwidth cap'}; no Microsoft access)`
			: 'OneDrive speed check (live Microsoft Graph)',
	);
	out.push(
		`${options.runs} run(s) per target${options.warmup ? ', first run discarded' : ''}. Token is static: acquisition is a function call, not a sign-in.`,
	);
	for (const { target, summary: m } of summaries) {
		out.push('', `== ${target.label} ==`, '');
		out.push(
			renderTable([
				[
					'run',
					'meta',
					'dl first byte',
					'dl total',
					'MB/s',
					'chunks',
					'client.open',
					'provider.open',
					'/content',
				],
				...target.runs.map((x, i) => [
					String(i + 1),
					fmtMs(x.meta.totalMs),
					fmtMs(x.download.firstByteMs),
					fmtMs(x.download.totalMs),
					fmtRate(x.download.mbps),
					String(x.download.chunks),
					x.library.error ? 'error' : fmtMs(x.library.totalMs),
					x.provider.error ? 'error' : fmtMs(x.provider.totalMs),
					x.content ? fmtMs(x.content.totalMs) : '-',
				]),
			]),
		);
		const row = (name, st, f = fmtMs) => [name, f(st.median), f(st.min), f(st.max)];
		out.push(
			'',
			renderTable([
				['metric', 'median', 'min', 'max'],
				row('token acquisition (static)', m.tokenMs),
				row('Graph metadata: time to first byte', m.metaHeadersMs),
				row('Graph metadata: total', m.metaTotalMs),
				row('Graph metadata: response size', m.metaBytes, fmtBytes),
				row('download: time to response headers', m.downloadHeadersMs),
				row('download: time to first byte', m.downloadFirstByteMs),
				row('download: total', m.downloadTotalMs),
				row('download: bytes', m.downloadBytes, fmtBytes),
				row('download: throughput', m.downloadMBps, fmtRate),
				row('download: chunks', m.downloadChunks, (v) =>
					Number.isFinite(v) ? String(Math.round(v)) : '-',
				),
				row('raw total (metadata + download)', m.rawMs),
				row('library client.open: total', m.libraryMs),
				row('  of which Graph response headers', m.libraryGraphHeadersMs),
				row('  of which download response headers', m.libraryDownloadHeadersMs),
				row('  time to first onProgress', m.libraryFirstProgressMs),
				row('library provider.open: total', m.providerMs),
				...(options.content
					? [
							row('/content: redirect (302)', m.contentRedirectMs),
							row('/content: redirect + download', m.contentTotalMs),
						]
					: []),
			]),
		);
		const errors = [
			...new Set(
				target.runs
					.flatMap((x) => [x.library.error, x.provider.error, x.content?.error])
					.filter(Boolean),
			),
		];
		if (errors.length) out.push('', `Notes: ${errors.join('; ')}`);
		out.push('', 'Reading:');
		for (const line of interpret({
			rawMeta: m.metaTotalMs.median,
			rawDownload: m.downloadTotalMs.median,
			rawTtfb: m.downloadFirstByteMs.median,
			library: m.libraryMs.median,
			provider: m.providerMs.median,
			content: m.contentTotalMs.median,
			bytes: m.downloadBytes.median,
		}))
			out.push(`  - ${line}`);
	}
	out.push('', '== listing (through the library) ==', '');
	out.push(
		renderTable([
			['call', 'entries', 'median', 'min', 'max'],
			...listing.map((l) => [
				l.name,
				String(l.entries),
				fmtMs(l.stats.median),
				fmtMs(l.stats.min),
				fmtMs(l.stats.max),
			]),
		]),
	);
	for (const l of listing) if (l.error) out.push(`Note: ${l.name} failed: ${l.error}`);
	console.log(out.join('\n'));
}

try {
	await main();
} catch (error) {
	// Never echo arbitrary errors: they can carry request details such as links or tokens.
	console.error(
		error instanceof SafeError ||
			error?.name === 'OneDriveError' ||
			/^(Invalid value|Use --help)/u.test(error?.message ?? '')
			? error.message
			: 'The speed check failed. Check the core build (bun run build), the token and the file access.',
	);
	process.exitCode = 1;
}
