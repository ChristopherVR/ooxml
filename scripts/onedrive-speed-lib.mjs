/** Pure helpers of onedrive-speed-check.mjs: statistics, formatting, encoding, arguments. */

export function median(values) {
	if (!values.length) return NaN;
	const sorted = [...values].sort((a, b) => a - b);
	const mid = sorted.length >> 1;
	return sorted.length % 2 ? sorted[mid] : (sorted[mid - 1] + sorted[mid]) / 2;
}

/** Median, min, max and mean of the finite numbers; NaN fields when there are none. */
export function stats(values) {
	const finite = values.filter((value) => Number.isFinite(value));
	if (!finite.length) return { n: 0, median: NaN, min: NaN, max: NaN, mean: NaN };
	return {
		n: finite.length,
		median: median(finite),
		min: Math.min(...finite),
		max: Math.max(...finite),
		mean: finite.reduce((sum, value) => sum + value, 0) / finite.length,
	};
}

export const fmtMs = (ms) =>
	!Number.isFinite(ms) ? '-' : ms >= 1000 ? `${(ms / 1000).toFixed(2)} s` : `${ms.toFixed(1)} ms`;
export const fmtBytes = (bytes) =>
	!Number.isFinite(bytes)
		? '-'
		: bytes >= 1048576
			? `${(bytes / 1048576).toFixed(2)} MB`
			: bytes >= 1024
				? `${(bytes / 1024).toFixed(1)} KB`
				: `${bytes} B`;
/** Megabytes (2^20 bytes) per second for `bytes` moved in `ms`. */
export const throughput = (bytes, ms) => (ms > 0 ? bytes / 1048576 / (ms / 1000) : NaN);
export const fmtRate = (mbps) => (Number.isFinite(mbps) ? `${mbps.toFixed(2)} MB/s` : '-');
export const percentOver = (value, base) =>
	base > 0 && Number.isFinite(value) ? ((value - base) / base) * 100 : NaN;

/** A drive path such as `Documents/My deck.pptx` as a Graph path (`/` kept between segments). */
export function encodeDrivePath(path) {
	const segments = path.split('/').filter((segment) => segment.length > 0);
	if (!segments.length) throw new Error('The OneDrive path is empty.');
	return segments.map((segment) => encodeURIComponent(segment)).join('/');
}

/** Same format as the library's entry ids: `!` inside each part is percent-encoded. */
const part = (value) => encodeURIComponent(value).replace(/!/gu, '%21');
export const entryId = (ref) => `${part(ref.driveId)}!${part(ref.itemId)}`;

/** Fixed-width text table; `rows` are arrays of strings, the first row is the header. */
export function renderTable(rows) {
	const widths = rows[0].map((_, column) =>
		Math.max(...rows.map((row) => String(row[column]).length)),
	);
	const line = (row) =>
		row
			.map((cell, i) =>
				i === 0 ? String(cell).padEnd(widths[i]) : String(cell).padStart(widths[i]),
			)
			.join('  ');
	return [
		line(rows[0]),
		widths.map((w) => '-'.repeat(w)).join('  '),
		...rows.slice(1).map(line),
	].join('\n');
}

export const SPEED_USAGE = `Usage: node scripts/onedrive-speed-check.mjs [options]

Measures where time goes when a file is downloaded through Microsoft Graph.

Live mode (environment variables only, nothing is read from arguments):
  OOXML_ONEDRIVE_ACCESS_TOKEN  short-lived delegated Graph token (required)
  OOXML_ONEDRIVE_SHARE_URL     a sharing link to an Office file (optional)
  OOXML_ONEDRIVE_PATH          a path in your own drive, e.g. Documents/deck.pptx (optional)
  OOXML_ONEDRIVE_FOLDER        a folder path in your own drive, to time listing (optional)
  At least one of SHARE_URL and PATH is needed. Tokens, links and download URLs are never printed.

Options:
  --runs=N         measured runs per target (default 5)
  --warmup         run once first and discard it (hides the cold connection cost)
  --no-content     skip the /content (302 redirect) comparison
  --json           print machine-readable results instead of the table
  --mock           no Microsoft access: measure a local server that imitates Graph
  --size=MB        mock file size in MB (default 10)
  --latency=ms     mock latency added to every response (default 30)
  --bandwidth=MBs  mock download host throughput cap in MB/s (default 0 = unlimited)
  --help           show this text`;

/** Parses the flags; throws a safe message for anything unknown or out of range. */
export function parseArgs(argv) {
	const options = {
		runs: 5,
		json: false,
		mock: false,
		warmup: false,
		content: true,
		help: false,
		size: 10,
		latency: 30,
		bandwidth: 0,
	};
	const number = (name, value, { min, integer }) => {
		const parsed = Number(value);
		if (
			value === '' ||
			!Number.isFinite(parsed) ||
			parsed < min ||
			(integer && !Number.isInteger(parsed))
		)
			throw new Error(`Invalid value for --${name}.`);
		return parsed;
	};
	for (const arg of argv) {
		const eq = arg.indexOf('=');
		const flag = eq < 0 ? arg : arg.slice(0, eq);
		const value = eq < 0 ? undefined : arg.slice(eq + 1);
		if (flag === '--help' && value === undefined) options.help = true;
		else if (flag === '--json' && value === undefined) options.json = true;
		else if (flag === '--mock' && value === undefined) options.mock = true;
		else if (flag === '--warmup' && value === undefined) options.warmup = true;
		else if (flag === '--no-content' && value === undefined) options.content = false;
		else if (flag === '--runs' && value !== undefined)
			options.runs = number('runs', value, { min: 1, integer: true });
		else if (flag === '--size' && value !== undefined)
			options.size = number('size', value, { min: 0.001 });
		else if (flag === '--latency' && value !== undefined)
			options.latency = number('latency', value, { min: 0 });
		else if (flag === '--bandwidth' && value !== undefined)
			options.bandwidth = number('bandwidth', value, { min: 0 });
		else throw new Error('Use --help for supported options.');
	}
	if (options.runs > 100) throw new Error('Invalid value for --runs.');
	return options;
}

const signed = (value) => `${value >= 0 ? '+' : '-'}${fmtMs(Math.abs(value))}`;

/** Plain-language reading of the median timings (ms; `bytes` is the download size). */
export function interpret({ rawMeta, rawDownload, rawTtfb, library, provider, content, bytes }) {
	const lines = [];
	const raw = rawMeta + rawDownload;
	if (Number.isFinite(raw) && raw > 0) {
		const share = (rawDownload / raw) * 100;
		lines.push(
			`The download host is ${share.toFixed(0)}% of the raw time (${fmtMs(rawDownload)} of ${fmtMs(raw)}); the Graph metadata call is ${fmtMs(rawMeta)}.`,
		);
		if (share >= 70)
			lines.push(
				`Time is dominated by the download host (${fmtRate(throughput(bytes, rawDownload))}); nothing in the client can make that part faster.`,
			);
		else if (rawMeta >= rawDownload)
			lines.push('Time is dominated by the Graph metadata round trip (latency, not size).');
		if (Number.isFinite(rawTtfb) && rawDownload > 0 && rawTtfb / rawDownload > 0.5)
			lines.push(
				`Over half of the download (${fmtMs(rawTtfb)}) passes before the first byte: latency, not bandwidth.`,
			);
	}
	if (Number.isFinite(library) && raw > 0) {
		const over = percentOver(library, raw);
		const delta = library - raw;
		const verdict =
			Math.abs(over) < 10 || Math.abs(delta) < 50
				? 'Within run-to-run noise: the library is not the bottleneck.'
				: over > 0
					? 'Worth investigating.'
					: 'The library is faster (connection reuse or a faster URL).';
		lines.push(
			`Library overhead: ${over >= 0 ? '+' : ''}${over.toFixed(1)}% (${signed(delta)}) over raw metadata + download. ${verdict}`,
		);
	}
	if (Number.isFinite(provider) && Number.isFinite(library) && library > 0)
		lines.push(
			`The file provider open is ${percentOver(provider, library).toFixed(1)}% versus client.open.`,
		);
	if (Number.isFinite(content) && raw > 0)
		lines.push(
			content < raw
				? `The /content redirect path is faster than metadata + downloadUrl (${fmtMs(content)} vs ${fmtMs(raw)}).`
				: `Metadata + downloadUrl is as fast or faster than /content (${fmtMs(raw)} vs ${fmtMs(content)}).`,
		);
	return lines;
}
