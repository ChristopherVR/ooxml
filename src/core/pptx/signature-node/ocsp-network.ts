/**
 * Guarded HTTP POST for OCSP responders. Responder URLs come from the
 * certificate's AIA extension, i.e. from the (untrusted) document, so the
 * request must not become an SSRF primitive.
 */

import dns from 'node:dns';
import net from 'node:net';

export interface OcspNetworkOptions {
	/** Allow responders on loopback, link-local, private and metadata addresses. Default false. */
	allowPrivateNetwork?: boolean;
}

export const OCSP_TIMEOUT_MS = 10_000;
export const OCSP_MAX_RESPONSE_BYTES = 1024 * 1024;

function isPrivateIpv4(ip: string): boolean {
	const p = ip.split('.').map(Number);
	const [a = 0, b = 0] = p;
	return (
		a === 0 ||
		a === 10 ||
		a === 127 ||
		(a === 169 && b === 254) ||
		(a === 172 && b >= 16 && b <= 31) ||
		(a === 192 && b === 168) ||
		(a === 100 && b >= 64 && b <= 127)
	);
}

/** True for loopback, link-local, RFC1918, CGNAT, unique-local, unspecified and metadata addresses. */
export function isPrivateAddress(address: string): boolean {
	const ip = address.replace(/^\[|\]$/g, '').split('%')[0] ?? '';
	if (net.isIPv4(ip)) {
		return isPrivateIpv4(ip);
	}
	if (!net.isIPv6(ip)) {
		return true;
	}
	const lower = ip.toLowerCase();
	const mapped = /^::ffff:(\d+\.\d+\.\d+\.\d+)$/.exec(lower);
	if (mapped?.[1]) {
		return isPrivateIpv4(mapped[1]);
	}
	const hex = /^::ffff:([0-9a-f]{1,4}):([0-9a-f]{1,4})$/.exec(lower);
	if (hex?.[1] && hex[2]) {
		const hi = Number.parseInt(hex[1], 16);
		const lo = Number.parseInt(hex[2], 16);
		return isPrivateIpv4(`${hi >> 8}.${hi & 255}.${lo >> 8}.${lo & 255}`);
	}
	if (lower === '::' || lower === '::1') {
		return true;
	}
	const first = Number.parseInt(lower.split(':')[0] || '0', 16);
	return (first & 0xfe00) === 0xfc00 || (first & 0xffc0) === 0xfe80;
}

async function assertPublicHost(hostname: string): Promise<void> {
	const host = hostname.replace(/^\[|\]$/g, '');
	const addresses = net.isIP(host)
		? [host]
		: (await dns.promises.lookup(host, { all: true })).map((entry) => entry.address);
	if (addresses.length === 0) {
		throw new Error('Responder host did not resolve');
	}
	for (const address of addresses) {
		if (isPrivateAddress(address)) {
			throw new Error('Responder host resolves to a private network address');
		}
	}
}

async function readCapped(response: Response, controller: AbortController): Promise<Buffer> {
	const declared = Number(response.headers.get('content-length') ?? 0);
	if (declared > OCSP_MAX_RESPONSE_BYTES) {
		controller.abort();
		throw new Error('Response too large');
	}
	if (!response.body) {
		return Buffer.alloc(0);
	}
	const reader = response.body.getReader();
	const chunks: Buffer[] = [];
	let total = 0;
	for (;;) {
		const { done, value } = await reader.read();
		if (done) break;
		total += value.byteLength;
		if (total > OCSP_MAX_RESPONSE_BYTES) {
			controller.abort();
			await reader.cancel().catch(() => undefined);
			throw new Error('Response too large');
		}
		chunks.push(Buffer.from(value));
	}
	return Buffer.concat(chunks);
}

export async function httpPost(
	url: string,
	body: Buffer,
	contentType: string,
	options: OcspNetworkOptions = {},
): Promise<Buffer> {
	let parsed: URL;
	try {
		parsed = new URL(url);
	} catch {
		throw new Error('Invalid URL');
	}
	if (parsed.protocol !== 'http:' && parsed.protocol !== 'https:') {
		throw new Error('Unsupported protocol');
	}
	if (!options.allowPrivateNetwork) {
		await assertPublicHost(parsed.hostname);
	}
	const controller = new AbortController();
	const timer = setTimeout(() => controller.abort(), OCSP_TIMEOUT_MS);
	try {
		const response = await fetch(parsed, {
			method: 'POST',
			headers: { 'Content-Type': contentType },
			body: new Uint8Array(body),
			redirect: 'manual',
			signal: controller.signal,
		});
		if (response.status >= 300 && response.status < 400) {
			throw new Error('Redirects are not followed');
		}
		return await readCapped(response, controller);
	} finally {
		clearTimeout(timer);
	}
}
