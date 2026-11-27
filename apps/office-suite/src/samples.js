/**
 * Opening a sample from the gallery (`samples/` on this site, scripts/build-samples-page.mjs):
 * `?sample=<id>` names an entry of `samples/manifest.json`. Nothing is fetched until the address
 * carries that parameter; then the manifest and that one file are read from this site, and a copy
 * joins the library like any opened file. The id only ever selects a manifest entry, so the
 * address cannot make the app fetch anything else. main.js loads this module on demand.
 */

export const SAMPLE_PARAM = 'sample';
const ID = /^[a-z0-9]+(?:-[a-z0-9]+)*$/u;
/** A sample file's path inside `samples/`: one product folder, one file. */
const FILE = /^[a-z]+\/[a-z0-9-]+\.(?:docx|xlsx|pptx|vsdx)$/u;
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
	return { id: ID.test(raw) ? raw : null, href: url.href };
}

/**
 * Look sample `id` up in the manifest of the site at `base`.
 * @param {string} id
 * @param {{ base: string | URL, fetch?: typeof fetch }} options
 * @returns {Promise<{ name: string, url: string, size: number }>}
 */
export async function findSample(id, { base, fetch = globalThis.fetch }) {
	const root = new URL('samples/', base);
	const listing = await fetch(new URL('manifest.json', root)).catch(() => undefined);
	if (!listing?.ok) throw new Error('The samples could not be loaded. Check your connection.');
	const manifest = await listing.json().catch(() => undefined);
	const entry = manifest?.samples?.find?.((sample) => sample?.id === id);
	if (!entry || typeof entry.file !== 'string' || !FILE.test(entry.file)) throw unavailable();
	return {
		name: entry.file.slice(entry.file.lastIndexOf('/') + 1),
		url: new URL(entry.file, root).href,
		size: Number(entry.size) || 0,
	};
}

/** The bytes of a sample `findSample` returned. */
export async function downloadSample(sample, { fetch = globalThis.fetch } = {}) {
	const response = await fetch(sample.url).catch(() => undefined);
	if (!response?.ok) throw unavailable();
	const bytes = new Uint8Array(await response.arrayBuffer());
	if (bytes.length === 0 || bytes.length > LIMIT) throw unavailable();
	return bytes;
}

/**
 * The library copy to reuse for `sample`: one fetched earlier that has not been edited since and
 * is still the published file's size, so opening a sample twice does not fill the library with
 * duplicates.
 */
export function untouchedCopy(documents, sample) {
	return documents.find(
		(doc) =>
			doc.sourceUrl === sample.url &&
			doc.revision === 1 &&
			!doc.cloud &&
			doc.bytes?.length === sample.size,
	);
}
