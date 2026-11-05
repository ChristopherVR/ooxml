/**
 * Local, opt-in OneDrive round trip against a real Microsoft Graph account. Never runs in CI: it
 * skips (exit 0) when no token is available.
 *
 * Usage:
 *   OOXML_ONEDRIVE_SHARE_URL='<sharing link to your own .pptx>' \
 *     node scripts/onedrive-live-roundtrip.mjs [--env <path to .env with GRAPH_TOKEN=...>]
 *       [--link-env <variable name>] [--restore]
 *
 * The token comes from the GRAPH_TOKEN (or OOXML_ONEDRIVE_ACCESS_TOKEN) environment variable, else
 * from GRAPH_TOKEN= in the --env file. The token and sharing link are never printed. Writes are
 * allowed (canWrite) because this is meant for your own file. Steps, each timed:
 *   1 open from the link; 2 parse with the pptx handler and edit A; 3 save through a session and
 *   re-download to verify A and the changed eTag; 4 two sessions from one version make
 *   non-overlapping edits B and C, the second hits the eTag conflict and merges; 5 optional
 *   --restore uploads the original bytes again (default keeps the edits).
 */
import { readFileSync } from 'node:fs';
import { pathToFileURL } from 'node:url';
import { checkToken, describeToken, normalizeToken } from './onedrive-token-check.mjs';
import { fmtBytes, fmtMs, fmtRate, renderTable, throughput } from './onedrive-speed-lib.mjs';

export function parseRoundtripArgs(argv) {
	const out = { env: undefined, linkEnv: 'OOXML_ONEDRIVE_SHARE_URL', restore: false };
	for (let i = 0; i < argv.length; i++) {
		const arg = argv[i];
		if (arg === '--restore') out.restore = true;
		else if (arg === '--env' || arg === '--link-env') {
			const value = argv[++i];
			if (!value) throw new Error(`${arg} needs a value`);
			if (arg === '--env') out.env = value;
			else out.linkEnv = value;
		} else throw new Error(`Unknown option: ${arg}`);
	}
	return out;
}

/** Read `KEY=value` lines; returns only the requested key. Never logs values. */
export function readEnvKey(text, key) {
	for (const line of String(text).split(/\r?\n/u)) {
		const match = /^\s*(?:export\s+)?([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*)$/u.exec(line);
		if (match && match[1] === key) return match[2].trim();
	}
	return undefined;
}

/** The normalized token from the environment or the .env file text, or '' when none. */
export function resolveToken(env, envFileText) {
	const raw =
		env.GRAPH_TOKEN ??
		env.OOXML_ONEDRIVE_ACCESS_TOKEN ??
		(envFileText === undefined ? undefined : readEnvKey(envFileText, 'GRAPH_TOKEN'));
	return normalizeToken(raw);
}

const MARK = 'ooxml live round-trip';
const SAFE_MESSAGE = /^(Token|Edit|The |Marker|No |Microsoft|Could|Merge)/u;

async function main() {
	const options = parseRoundtripArgs(process.argv.slice(2));
	let fileText;
	if (options.env) {
		try {
			fileText = readFileSync(options.env, 'utf8');
		} catch {
			fileText = undefined;
		}
	}
	const token = resolveToken(process.env, fileText);
	const shareUrl = process.env[options.linkEnv];
	if (!token) {
		console.log('SKIP: no GRAPH_TOKEN available (environment or --env file). Nothing was run.');
		return 0;
	}
	if (!shareUrl) {
		console.log(`SKIP: set ${options.linkEnv} to a sharing link. Nothing was run.`);
		return 0;
	}
	const rows = [];
	let step = 'token check';
	const timed = async (name, fn) => {
		step = name;
		const t = performance.now();
		const value = await fn();
		return { value, ms: performance.now() - t };
	};
	const add = (name, ms, note = '') => rows.push([name, fmtMs(ms), note]);
	const table = () => renderTable([['Step', 'Time', 'Detail'], ...rows]);
	try {
		const { notes } = describeToken(token);
		if (notes.some((n) => /expired|another resource|authorization code/u.test(n)))
			throw new Error(`Token unusable: ${notes.join('; ')}`);
		await checkToken(token);
		const lib = await import('../src/core/dist/embedding/onedrive/index.mjs');
		const pptx = await import('../src/core/dist/pptx/index.mjs');
		// Every request, by kind only (never its URL): where the time of a save goes.
		const log = [];
		const traced = async (input, init) => {
			const url = new URL(String(input instanceof Request ? input.url : input));
			const kind = url.pathname.endsWith('/createUploadSession')
				? 'create upload session'
				: url.hostname === 'graph.microsoft.com'
					? `graph ${init?.method ?? 'GET'}`
					: init?.method === 'PUT'
						? 'upload fragment'
						: init?.method === 'DELETE'
							? 'cancel upload'
							: 'download';
			const t = performance.now();
			const response = await fetch(input, init);
			log.push({
				kind,
				ms: performance.now() - t,
				status: response.status,
				sent: init?.body?.length ?? 0,
			});
			return response;
		};
		const requests = (from) => {
			const parts = new Map();
			for (const entry of log.slice(from)) {
				const part = parts.get(entry.kind) ?? { n: 0, ms: 0, statuses: new Set() };
				part.n++;
				part.ms += entry.ms;
				part.statuses.add(entry.status);
				parts.set(entry.kind, part);
			}
			return [...parts]
				.map(([kind, p]) => `${kind} x${p.n} ${fmtMs(p.ms)} (${[...p.statuses].join('/')})`)
				.join('; ');
		};
		const client = lib.createOneDriveClient({
			getAccessToken: async () => token,
			canWrite: () => true,
			fetch: traced,
		});
		const mergeStats = { calls: 0, ms: 0 };
		const sessionOptions = {
			merge: async ({ base, local, remote }) => {
				mergeStats.calls++;
				const t = performance.now();
				const result = await pptx.mergePptxSnapshots(base.bytes, local, remote.bytes);
				mergeStats.ms += performance.now() - t;
				return result;
			},
		};
		const text = async (bytes) =>
			JSON.stringify((await pptx.inspectPresentationText(bytes)).slides);
		const edit = async (bytes, target, marker) => {
			const handler = new pptx.PptxHandler();
			const data = await handler.load(new Uint8Array(bytes).buffer);
			const candidates = [];
			data.slides.forEach((slide, si) =>
				slide.elements.forEach((el) => {
					if (el.textSegments?.length) candidates.push({ si, el });
				}),
			);
			if (!candidates.length) throw new Error('The deck has no editable text runs.');
			const pick =
				target === 'first'
					? candidates[0]
					: ([...candidates].reverse().find((c) => c.el !== candidates[0].el) ?? candidates[0]);
			const seg = pick.el.textSegments[pick.el.textSegments.length - 1];
			seg.text = `${seg.text} [${marker}]`;
			pick.el.text = pick.el.textSegments.map((s) => s.text).join('');
			const out = await handler.save(data.slides);
			handler.dispose?.();
			return { bytes: out, slide: pick.si + 1 };
		};
		const stamp = new Date().toISOString();

		const opened = await timed('1 open', () => client.open(shareUrl));
		const original = opened.value.bytes;
		add(
			'1 open (metadata+download)',
			opened.ms,
			`${opened.value.file.name}, ${fmtBytes(original.length)}, ${fmtRate(throughput(original.length, opened.ms))}`,
		);

		const markerA = `${MARK} A ${stamp}`;
		const parse = await timed('2 parse', () => pptx.inspectPresentationText(original));
		add('2a parse (handler load)', parse.ms, `${parse.value.slides.length} slides`);
		const a = await timed('2 edit A', () => edit(original, 'first', markerA));
		add('2b edit A + serialize', a.ms, `${fmtBytes(a.value.bytes.length)}, slide ${a.value.slide}`);

		const s0 = lib.createOneDriveSession(client, opened.value, sessionOptions);
		let mark = log.length;
		const phases = [];
		const onSaveProgress = (p) => {
			if (phases.at(-1) !== p.phase) phases.push(p.phase);
		};
		const saveA = await timed('3 save A', () => s0.save(a.value.bytes, { onSaveProgress }));
		add('3a upload A', saveA.ms, fmtRate(throughput(a.value.bytes.length, saveA.ms)));
		console.log(`  3a requests: ${requests(mark)}; progress: ${phases.join(' > ')}`);

		// The everyday case: the same person saves again in the same session. It must not conflict.
		const markerA2 = `${MARK} A2 ${stamp}`;
		const a2 = await timed('3 edit A2', () => edit(a.value.bytes, 'first', markerA2));
		mark = log.length;
		phases.length = 0;
		const mergesBefore = mergeStats.calls;
		const saveA2 = await timed('3 save A2 (same session)', () =>
			s0.save(a2.value.bytes, { onSaveProgress }),
		);
		const conflicts = log.slice(mark).filter((e) => e.status === 412 || e.status === 409).length;
		add(
			'3c second save, same session',
			saveA2.ms,
			conflicts
				? `${conflicts} revision conflict(s), ${mergeStats.calls - mergesBefore} merge(s)`
				: 'no conflict',
		);
		console.log(`  3c requests: ${requests(mark)}; progress: ${phases.join(' > ')}`);
		if (mergeStats.calls > mergesBefore)
			throw new Error('A second save by the only editor needed a merge.');
		const fresh = await timed('3 re-download', () => client.open(shareUrl));
		add('3b re-download', fresh.ms, fmtBytes(fresh.value.bytes.length));
		if (!(await text(fresh.value.bytes)).includes(markerA2))
			throw new Error('Edit A2 is missing after re-download.');
		if (fresh.value.file.eTag === opened.value.file.eTag)
			throw new Error('The eTag did not change after saving.');
		console.log('Verified: edit A present, eTag changed.');

		const base = fresh.value;
		const s1 = lib.createOneDriveSession(client, base, sessionOptions);
		const s2 = lib.createOneDriveSession(client, base, sessionOptions);
		const markerB = `${MARK} B ${stamp}`;
		const markerC = `${MARK} C ${stamp}`;
		const b = await timed('4 edit B', () => edit(base.bytes, 'first', markerB));
		const c = await timed('4 edit C', () => edit(base.bytes, 'other', markerC));
		add('4a edit B / C', b.ms + c.ms, `B slide ${b.value.slide}, C slide ${c.value.slide}`);
		mark = log.length;
		const saveB = await timed('4 save B', () => s1.save(b.value.bytes));
		add('4b session 1 save B', saveB.ms);
		console.log(`  4b requests: ${requests(mark)}`);
		mark = log.length;
		let saveC;
		try {
			saveC = await timed('4 save C (conflict, merge)', () => s2.save(c.value.bytes));
		} catch (error) {
			if (error instanceof lib.OneDriveConflictError)
				console.error(
					`Merge unresolved (${mergeStats.calls} attempts): ${(error.conflict.issues ?? []).slice(0, 10).join(' | ')}`,
				);
			throw error;
		}
		console.log(`  4c requests: ${requests(mark)}`);
		add(
			'4c session 2 save C (conflict+merge)',
			saveC.ms,
			`merge ${fmtMs(mergeStats.ms)}, merge attempts ${mergeStats.calls}`,
		);
		const final = await timed('4 verify', () => client.open(shareUrl));
		add('4d re-download', final.ms, fmtBytes(final.value.bytes.length));
		const finalText = await text(final.value.bytes);
		for (const [name, m] of [
			['A', markerA],
			['B', markerB],
			['C', markerC],
		])
			if (!finalText.includes(m)) throw new Error(`Marker ${name} is missing after merge.`);
		if (mergeStats.calls < 1)
			throw new Error('No conflict occurred, so the merge path was not exercised.');
		console.log('Verified: A, B and C all present after the merge.');

		if (options.restore) {
			const sr = lib.createOneDriveSession(client, final.value, sessionOptions);
			const r = await timed('5 restore', () => sr.save(original));
			add('5 restore original', r.ms);
			console.log('Restored the original content.');
		} else console.log('Kept the edits (use --restore to put the original content back).');
		console.log(table());
		return 0;
	} catch (error) {
		console.log(table());
		const code = error?.code ? ` code=${error.code}` : '';
		const status = error?.status ? ` status=${error.status}` : '';
		const named = error?.name === 'OneDriveError' || error?.name === 'OneDriveConflictError';
		const message =
			named || SAFE_MESSAGE.test(String(error?.message))
				? String(error.message).slice(0, 400)
				: `unexpected ${error?.name ?? 'error'}`;
		console.error(`FAILED at step "${step}":${code}${status} ${message}`);
		return 1;
	}
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href)
	process.exitCode = await main();
