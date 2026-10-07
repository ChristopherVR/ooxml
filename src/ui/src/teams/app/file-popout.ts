import { contentUrl, sanitizeAttachment } from 'ooxml-core/teams';
import type { OpenFileDetail } from './teams-app';

const MARKER = 'openteams-file';
/** A separate tab runs the same UI bundle and reuses the native content renderer. */
export function filePopoutUrl(detail: OpenFileDetail, pageUrl: string): string | null {
	const url = contentUrl(detail.url, pageUrl);
	const attachment = sanitizeAttachment(detail.attachment);
	if (!url || !attachment) return null;
	const page = new URL(pageUrl);
	page.searchParams.set(MARKER, '1');
	page.hash = encodeURIComponent(JSON.stringify({ attachment, url }));
	if (page.hash.length > 32_768) return null;
	return page.href;
}
export function filePopoutDetail(pageUrl: string): OpenFileDetail | null {
	try {
		const page = new URL(pageUrl);
		if (page.searchParams.get(MARKER) !== '1' || page.hash.length > 32_768) return null;
		const raw = JSON.parse(decodeURIComponent(page.hash.slice(1))) as Record<string, unknown>;
		const attachment = sanitizeAttachment(raw.attachment);
		const url = contentUrl(typeof raw.url === 'string' ? raw.url : undefined, pageUrl);
		return attachment && url ? { attachment, url } : null;
	} catch {
		return null;
	}
}
