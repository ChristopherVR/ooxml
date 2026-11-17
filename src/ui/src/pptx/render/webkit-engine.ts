/**
 * Engine detection for the WebKit workarounds in export and print. Feature detection cannot tell
 * these behaviours apart (iOS prints the parent page for an iframe's `print()`, decodes images
 * inside an SVG image lazily, and caps canvas area), so the user agent is read, and only to pick a
 * workaround that is harmless elsewhere.
 */

/** The `navigator` fields the detection reads, injectable for tests. */
export interface WebKitEnvironment {
	userAgent?: string;
	maxTouchPoints?: number;
}

/** Whether the browser is iOS or iPadOS WebKit (every browser there, Safari or not). */
export function isIosWebKit(env: WebKitEnvironment | undefined): boolean {
	const userAgent = env?.userAgent ?? '';
	if (/iPad|iPhone|iPod/.test(userAgent)) return true;
	// iPadOS 13 and later report a desktop Mac user agent; only the touch screen tells them apart.
	return /Macintosh/.test(userAgent) && (env?.maxTouchPoints ?? 0) > 1;
}

/** Whether the page runs on WebKit: Safari on macOS, or any browser on iOS and iPadOS. */
export function isWebKitEngine(env: WebKitEnvironment | undefined): boolean {
	if (isIosWebKit(env)) return true;
	const userAgent = env?.userAgent ?? '';
	return (
		/AppleWebKit/.test(userAgent) &&
		/Safari\//.test(userAgent) &&
		!/Chrome|Chromium|Edg\/|Android/.test(userAgent)
	);
}

/** `globalThis.navigator`, or `undefined` outside a browser. */
export function currentWebKitEnvironment(): WebKitEnvironment | undefined {
	return typeof navigator === 'undefined' ? undefined : navigator;
}
