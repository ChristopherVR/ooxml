/**
 * Opening a sample from the gallery (`samples/` on this site, scripts/build-samples-page.mjs):
 * `?sample=<id>` names an entry of `samples/manifest.json`. Nothing is fetched until the address
 * carries that parameter; then the manifest and that one file are read from this site, and a copy
 * joins the library like any opened file. The id only ever selects a manifest entry, so the
 * address cannot make the app fetch anything else, and the file must be the size and SHA-256 the
 * manifest records. main.js loads this module on demand.
 */

export const SAMPLE_PARAM = 'sample';
/** A sample id: a lower-case slug. The gallery build refuses a manifest entry that is not one. */
export const SAMPLE_ID = /^[a-z0-9]+(?:-[a-z0-9]+)*$/u;
/** A sample file's path inside `samples/`: one product folder, one file. */
const FILE = /^[a-z]+\/[a-z0-9-]+\.(?:docx|xlsx|pptx|vsdx)$/u;
const SHA256 = /^[0-9a-f]{64}$/u;
const LIMIT = 32 * 1024 * 1024;
const unavailable = () => new Error('That sample is not available.');

/**
 * Read the sample id from `href` and the address without it (so a reload does not open another
 * copy).
 * @returns {{ id: string | null, href: string } | null} null when the address names no sample;
 * `id` is null when it names one that cannot exist
 */
export function takeSampleLink(href) {
	const url = new URL(href);
	const raw = url.searchParams.get(SAMPLE_PARAM);
	if (raw === null) return null;
	url.searchParams.delete(SAMPLE_PARAM);
	return { id: SAMPLE_ID.test(raw) ? raw : null, href: url.href };
}

/**
 * Look sample `id` up in the manifest of the site at `base`.
 * @param {string} id
 * @param {{ base: string | URL, fetch?: typeof fetch }} options
 * @returns {Promise<{ name: string, url: string, size: number, sha256: string }>}
 */
export async function findSample(id, { base, fetch = globalThis.fetch }) {
	const root = new URL('samples/', base);
	const listing = await fetch(new URL('manifest.json', root)).catch(() => undefined);
	if (!listing?.ok) throw new Error('The samples could not be loaded. Check your connection.');
	const manifest = await listing.json().catch(() => undefined);
	const entry = manifest?.samples?.find?.((sample) => sample?.id === id);
	if (
		!entry ||
		typeof entry.file !== 'string' ||
		!FILE.test(entry.file) ||
		!Number.isSafeInteger(entry.size) ||
		entry.size <= 0 ||
		entry.size > LIMIT ||
		typeof entry.sha256 !== 'string' ||
		!SHA256.test(entry.sha256)
	)
		throw unavailable();
	return {
		name: entry.file.slice(entry.file.lastIndexOf('/') + 1),
		url: new URL(entry.file, root).href,
		size: entry.size,
		sha256: entry.sha256,
	};
}

/**
 * The SHA-256 of `bytes` as lower-case hex, or undefined where the page has no Web Crypto (a
 * deployment served over plain http): there the size alone is compared.
 */
export async function sha256Hex(bytes, subtle = globalThis.crypto?.subtle) {
	if (!subtle) return undefined;
	const digest = new Uint8Array(await subtle.digest('SHA-256', bytes));
	return Array.from(digest, (byte) => byte.toString(16).padStart(2, '0')).join('');
}

/** Whether `bytes` are the file `sample` describes: its size and, with Web Crypto, its hash. */
async function matches(bytes, sample, subtle) {
	if (bytes?.length !== sample.size) return false;
	const hash = await sha256Hex(bytes, subtle);
	return hash === undefined || hash === sample.sha256;
}

/**
 * The bytes of a sample `findSample` returned, refused unless they are the file the manifest
 * describes (a stale cached copy, or a file swapped on the server, is not opened).
 */
export async function downloadSample(
	sample,
	{ fetch = globalThis.fetch, subtle = globalThis.crypto?.subtle } = {},
) {
	const response = await fetch(sample.url).catch(() => undefined);
	if (!response?.ok) throw unavailable();
	const bytes = new Uint8Array(await response.arrayBuffer());
	if (!(await matches(bytes, sample, subtle))) throw unavailable();
	return bytes;
}

/**
 * The library copy to reuse for `sample`: one that has never been edited and still holds exactly
 * the published bytes, so opening a sample twice does not fill the library with duplicates.
 */
export async function untouchedCopy(
	documents,
	sample,
	{ subtle = globalThis.crypto?.subtle } = {},
) {
	for (const doc of documents) {
		if (doc.revision !== 1 || doc.cloud || doc.parent) continue;
		if (doc.bytes?.length !== sample.size) continue;
		const hash = await sha256Hex(doc.bytes, subtle);
		// Without Web Crypto, only a copy fetched from the sample's own address is taken for it.
		if (hash === undefined ? doc.sourceUrl === sample.url : hash === sample.sha256) return doc;
	}
	return undefined;
}

/**
 * A visible notice above the tabs when a sample cannot be opened, with the way back to the
 * gallery; the status line announces the same message (ui.js `notify`), so this adds no second
 * live region. Returns the notice element.
 * @param {string} message
 * @param {{ base: string | URL, host: Element, document?: Document }} options
 */
export function showSampleNotice(message, { base, host, document: doc = globalThis.document }) {
	doc.getElementById('sample-notice')?.remove();
	const notice = doc.createElement('div');
	notice.id = 'sample-notice';
	notice.className = 'update-banner';
	const text = doc.createElement('span');
	text.textContent = message;
	const browse = doc.createElement('a');
	browse.href = new URL('samples/', base).href;
	browse.textContent = 'Browse samples';
	const close = doc.createElement('button');
	close.textContent = 'Dismiss';
	close.onclick = () => notice.remove();
	notice.append(text, browse, close);
	host.prepend(notice);
	return notice;
}
