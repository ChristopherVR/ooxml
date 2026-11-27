import { resolveWebVideo } from 'ooxml-core/docx';

/** One video a search found: its player or watch address, which the allow-list must accept. */
export interface OnlineVideoSearchResult {
	title: string;
	/** An https embed or watch address; results whose host `onlineVideoHosts` lacks are not offered. */
	url: string;
	/** An https picture of the video, loaded by the browser from the address the provider gives. */
	thumbnailUrl?: string;
	/** Who published the video, shown under its title. */
	publisher?: string;
	durationSeconds?: number;
}

/**
 * A host's video search for Insert > Online Video (bring your own; none by default). The editor
 * sends only the words the user types in the search box, when they press Search.
 */
export interface OnlineVideoSearchProvider {
	/** The service named beside the search box ("Search YouTube"). */
	readonly name: string;
	search(
		query: string,
		options: { signal: AbortSignal; locale: string },
	): Promise<readonly OnlineVideoSearchResult[]>;
}

/** The longest query the editor sends and the most results it lists. */
export const MAX_VIDEO_QUERY_LENGTH = 200;
export const MAX_VIDEO_RESULTS = 25;

/** The results a user may insert: an https address on an allowed host, capped and de-duplicated. */
export function allowedVideoResults(
	results: readonly OnlineVideoSearchResult[],
	hosts: readonly string[],
): OnlineVideoSearchResult[] {
	const seen = new Set<string>();
	const allowed: OnlineVideoSearchResult[] = [];
	for (const result of results) {
		if (allowed.length >= MAX_VIDEO_RESULTS) break;
		if (typeof result?.url !== 'string' || typeof result.title !== 'string') continue;
		const resolved = resolveWebVideo(result.url, hosts);
		if (!resolved.ok || seen.has(resolved.spec.url)) continue;
		seen.add(resolved.spec.url);
		const thumbnail = safeThumbnail(result.thumbnailUrl);
		allowed.push({
			title: result.title.slice(0, 300),
			url: resolved.spec.url,
			...(thumbnail ? { thumbnailUrl: thumbnail } : {}),
			...(typeof result.publisher === 'string'
				? { publisher: result.publisher.slice(0, 120) }
				: {}),
			...(Number.isFinite(result.durationSeconds) && result.durationSeconds! >= 0
				? { durationSeconds: Math.round(result.durationSeconds!) }
				: {}),
		});
	}
	return allowed;
}

function safeThumbnail(url: string | undefined): string | undefined {
	if (typeof url !== 'string') return undefined;
	try {
		const parsed = new URL(url);
		return parsed.protocol === 'https:' && !parsed.username && !parsed.password
			? parsed.href
			: undefined;
	} catch {
		return undefined;
	}
}

/** `m:ss` or `h:mm:ss`. */
export function formatVideoDuration(seconds: number): string {
	const h = Math.floor(seconds / 3600);
	const m = Math.floor((seconds % 3600) / 60);
	const s = seconds % 60;
	const two = (value: number) => String(value).padStart(2, '0');
	return h ? `${h}:${two(m)}:${two(s)}` : `${m}:${two(s)}`;
}
