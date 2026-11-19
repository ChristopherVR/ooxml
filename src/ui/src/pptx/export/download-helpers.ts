/**
 * Browser download helpers shared by every binding's export pipeline. These are
 * the only DOM-touching helpers in the shared export subtree: they create a
 * transient `<a download>`, click it, and revoke the object URL after a short
 * delay. {@link downloadBlob} and {@link downloadDataUrl} run every filename
 * through the `sanitizeDownloadFilename` guard (null/traversal/length)
 * themselves, so no binding can accidentally skip it.
 *
 * The standard revoke delay is 200ms: long enough for the browser to begin the
 * download, short enough not to leak the object URL.
 */

import { isBrowserOpenableDataUrl } from 'ooxml-core/opc';

/** Default fallback name used when a download filename is empty or unusable. */
const FALLBACK_DOWNLOAD_NAME = 'presentation.pptx',
	/** Fallback stem used by {@link resolveExportBaseName} when no usable name remains. */
	EXPORT_BASE_NAME_FALLBACK = 'presentation',
	/**
	 * Extensions {@link resolveExportBaseName} strips by default: every format the
	 * export surface (PNG/PDF/GIF/WebM) or the original source deck (pptx) can be
	 * named after. Bare, lower-case, no leading dot.
	 */
	DEFAULT_EXPORT_STRIP_EXTENSIONS = ['pptx', 'pdf', 'png', 'gif', 'webm'],
	/** Revoke delay (ms) after triggering a download click. */
	REVOKE_DELAY_MS = 200,
	/**
	 * Revoke delay (ms) after opening a payload in a new tab. Longer than the
	 * download delay: the new document may still be fetching the object URL (a PDF
	 * viewer, an image) well after the current task finishes, so we keep the URL
	 * alive for a minute before releasing it.
	 */
	OPEN_REVOKE_DELAY_MS = 60_000;

/**
 * Strip control characters, filesystem-reserved characters, and path-traversal
 * sequences from a user-supplied download filename. CR/LF in particular can
 * corrupt `Content-Disposition` headers when a server-side proxy re-emits the
 * download name; leading dots can produce hidden files; the rest are simply
 * disallowed by Windows or unsafe to render in UI.
 *
 * Empty / whitespace-only input (or a non-string) falls back to
 * `presentation.pptx`. Over-long names are truncated to 200 chars, preserving a
 * short trailing extension when present.
 *
 * @param input - The raw, possibly hostile filename.
 * @returns A filesystem-safe download name (never empty).
 */
export function sanitizeDownloadFilename(input: string | undefined | null): string {
	if (typeof input !== 'string' || input.trim().length === 0) {
		return FALLBACK_DOWNLOAD_NAME;
	}
	let cleaned = input
		// eslint-disable-next-line no-control-regex
		.replace(/[\x00-\x1f\x7f"\\/:*?<>|]/gu, '_')
		.replace(/\.\./gu, '__')
		.replace(/^\.+/u, '')
		.trim();
	if (cleaned.length === 0) {
		return FALLBACK_DOWNLOAD_NAME;
	}
	if (cleaned.length > 200) {
		// Preserve the extension when truncating.
		const dot = cleaned.lastIndexOf('.');
		if (dot > 0 && cleaned.length - dot <= 16) {
			const ext = cleaned.slice(dot);
			cleaned = cleaned.slice(0, 200 - ext.length) + ext;
		} else {
			cleaned = cleaned.slice(0, 200);
		}
	}
	return cleaned;
}

/**
 * Derive the base (extensionless) file name for an export download from a
 * user-supplied source file name: trims whitespace, strips one trailing
 * extension from `extensions` (case-insensitive), and falls back to
 * `presentation` when `sourceName` is `undefined` or the result is empty.
 *
 * Used by every binding's export controller to turn a loaded deck's source
 * name (e.g. `Deck.pptx`) into the stem for its own export downloads (e.g.
 * `Deck-slide-1.png`, `Deck.gif`) without doubling up the extension.
 *
 * @param sourceName - The raw source file name, or `undefined` when none is
 *   known yet.
 * @param extensions - Extensions to strip, without the leading dot. Defaults
 *   to the full export-surface set (`pptx`, `pdf`, `png`, `gif`, `webm`).
 */
export function resolveExportBaseName(
	sourceName: string | undefined,
	extensions: readonly string[] = DEFAULT_EXPORT_STRIP_EXTENSIONS,
): string {
	if (sourceName === undefined) {
		return EXPORT_BASE_NAME_FALLBACK;
	}
	const pattern = new RegExp(`\\.(?:${extensions.join('|')})$`, 'iu'),
		trimmed = sourceName.trim().replace(pattern, '');
	return trimmed === '' ? EXPORT_BASE_NAME_FALLBACK : trimmed;
}

/**
 * Trigger a browser download for a Blob. `filename` is run through
 * {@link sanitizeDownloadFilename} before being applied, so every binding gets
 * a safe `Content-Disposition`/on-disk name regardless of whether it sanitized
 * the name itself: a source deck's own (attacker-controllable) file name can
 * reach this call unsanitized via `resolveExportBaseName`, so this is the one
 * place the guard cannot be skipped.
 *
 * @param blob     - The content to download.
 * @param filename - The suggested download name (sanitized before use).
 */
export function downloadBlob(blob: Blob, filename: string): void {
	const url = URL.createObjectURL(blob),
		anchor = document.createElement('a');
	anchor.href = url;
	anchor.download = sanitizeDownloadFilename(filename);
	document.body.appendChild(anchor);
	anchor.click();
	// Defer cleanup so the browser has time to start the download.
	setTimeout(() => {
		anchor.remove();
		URL.revokeObjectURL(url);
	}, REVOKE_DELAY_MS);
}

/**
 * Trigger a browser download for a data-URL string. `filename` is run through
 * {@link sanitizeDownloadFilename} before being applied; see {@link downloadBlob}.
 *
 * Anything but a `data:` or `blob:` URL is ignored: a link with another scheme,
 * such as `javascript:`, would be followed instead of downloaded.
 *
 * @param dataUrl  - The `data:` (or object) URL to download.
 * @param filename - The suggested download name (sanitized before use).
 */
export function downloadDataUrl(dataUrl: string, filename: string): void {
	if (!/^(?:data|blob):/iu.test(dataUrl)) {
		return;
	}
	const anchor = document.createElement('a');
	anchor.href = dataUrl;
	anchor.download = sanitizeDownloadFilename(filename);
	document.body.appendChild(anchor);
	anchor.click();
	setTimeout(() => {
		anchor.remove();
	}, REVOKE_DELAY_MS);
}

/**
 * Convert a `data:` URL into a {@link Blob}, preserving its MIME type. Handles
 * both base64 and percent-encoded payloads. Returns `undefined` for a non-data
 * URL or a payload that cannot be decoded, so callers can fall back gracefully.
 *
 * @param dataUrl - A `data:` URL string.
 */
export function dataUrlToBlob(dataUrl: string): Blob | undefined {
	const match = /^data:(?<mime>[^;,]*)(?<base64>;base64)?,(?<payload>[\s\S]*)$/u.exec(dataUrl);
	if (!match?.groups) {
		return undefined;
	}
	// eslint-disable-next-line one-var -- separated from `match` above by a guard clause
	const mime = match.groups.mime || 'application/octet-stream',
		isBase64 = Boolean(match.groups.base64),
		payload = match.groups.payload ?? '';
	try {
		if (isBase64) {
			const binary = atob(payload),
				bytes = new Uint8Array(binary.length);
			for (let i = 0; i < binary.length; i++) {
				bytes[i] = binary.charCodeAt(i);
			}
			return new Blob([bytes], { type: mime });
		}
		return new Blob([decodeURIComponent(payload)], { type: mime });
	} catch {
		return undefined;
	}
}

/**
 * Open an embedded payload in a new browser tab. The tab shares this page's
 * origin, so only a `data:` URL whose own type is inert in a browser (PDF,
 * plain text, CSV, JSON, raster images; see `isBrowserOpenableDataUrl` in
 * `ooxml-core/opc`) is opened. Any other URL or type is refused, whatever a
 * separate MIME field claims. Chromium (and other browsers) silently refuse to
 * navigate a new top-level browsing context straight to a `data:` URL, so the
 * payload is converted to a Blob object URL first and revoked after
 * {@link OPEN_REVOKE_DELAY_MS}, once the new document has had time to fetch it.
 *
 * @param url - The payload URL (a recovered `data:` URL).
 * @returns Whether a tab was opened; a caller offers a download otherwise.
 */
export function openUrlInNewTab(url: string): boolean {
	const blob = isBrowserOpenableDataUrl(url) ? dataUrlToBlob(url) : undefined;
	if (!blob) {
		return false;
	}
	const target = URL.createObjectURL(blob),
		// NB: no `noopener` here. A `blob:` object URL is resolved from the opener's
		// origin-partitioned blob store, and Chromium refuses to resolve it in the
		// disconnected browsing context `noopener` creates (the new tab lands on an
		// empty document). We instead sever the child's back-reference to us
		// afterwards, which mitigates reverse-tabnabbing without breaking the blob.
		opened = window.open(target, '_blank');
	if (opened) {
		try {
			opened.opener = null;
		} catch {
			// Some browsers disallow reassigning `opener`; best-effort only.
		}
	}
	setTimeout(() => {
		URL.revokeObjectURL(target);
	}, OPEN_REVOKE_DELAY_MS);
	return opened !== null;
}

/**
 * Show an embedded payload the way a browser can: in a new tab when
 * {@link openUrlInNewTab} allows it and `download` is false, otherwise as a
 * file download. Every binding's OLE verb runs through this, so a payload no
 * tab may open still reaches the user, as a file.
 */
export function openOrDownloadPayload(url: string, fileName: string, download = false): void {
	if (!download && openUrlInNewTab(url)) {
		return;
	}
	downloadDataUrl(url, fileName);
}
