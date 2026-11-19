/**
 * Guarded HTTP POST for OCSP responders. Responder URLs come from the
 * certificate's AIA extension, i.e. from the (untrusted) document, so the
 * request must not become an SSRF primitive: it goes only to globally routable
 * addresses, and the socket connects to the very addresses that were checked.
 */

import dns from 'node:dns';
import http, { type IncomingMessage } from 'node:http';
import https from 'node:https';
import net from 'node:net';

export interface OcspNetworkOptions {
	/** Allow responders on addresses that are not globally routable. Default false. */
	allowPrivateNetwork?: boolean;
}

export const OCSP_TIMEOUT_MS = 10_000;
export const OCSP_MAX_RESPONSE_BYTES = 1024 * 1024;

type Subnet = readonly [address: string, prefix: number];

function blockListOf(family: 'ipv4' | 'ipv6', subnets: readonly Subnet[]): net.BlockList {
	const list = new net.BlockList();
	for (const [address, prefix] of subnets) list.addSubnet(address, prefix, family);
	return list;
}

/** IANA special-purpose IPv4 ranges that are not globally reachable, multicast and reserved. */
const IPV4_NOT_GLOBAL = blockListOf('ipv4', [
	['0.0.0.0', 8],
	['10.0.0.0', 8],
	['100.64.0.0', 10],
	['127.0.0.0', 8],
	['169.254.0.0', 16],
	['172.16.0.0', 12],
	['192.0.0.0', 24],
	['192.0.2.0', 24],
	['192.88.99.0', 24],
	['192.168.0.0', 16],
	['198.18.0.0', 15],
	['198.51.100.0', 24],
	['203.0.113.0', 24],
	['224.0.0.0', 4],
	['240.0.0.0', 4],
]);

/** Global unicast, the only IPv6 space a public responder is in. */
const IPV6_GLOBAL = blockListOf('ipv6', [['2000::', 3]]);

/** The parts of global unicast set aside: IETF protocols, documentation and 6to4 tunnels. */
const IPV6_NOT_GLOBAL = blockListOf('ipv6', [
	['2001::', 23],
	['2001:db8::', 32],
	['2002::', 16],
	['3fff::', 20],
]);

/** An IPv4-mapped or NAT64 address in canonical form: its last 32 bits are the real target. */
const EMBEDDED_IPV4 = /^(?:::ffff|64:ff9b:):([0-9a-f]{1,4}):([0-9a-f]{1,4})$/;

/** True for every address that is not globally routable, and for anything not an address. */
export function isPrivateAddress(address: string): boolean {
	const ip = address.replace(/^\[|\]$/g, '').split('%')[0] ?? '';
	if (net.isIPv4(ip)) {
		return IPV4_NOT_GLOBAL.check(ip, 'ipv4');
	}
	if (!net.isIPv6(ip)) {
		return true;
	}
	const canonical = new URL(`http://[${ip}]/`).hostname.slice(1, -1);
	const embedded = EMBEDDED_IPV4.exec(canonical);
	if (embedded?.[1] && embedded[2]) {
		const high = Number.parseInt(embedded[1], 16);
		const low = Number.parseInt(embedded[2], 16);
		return IPV4_NOT_GLOBAL.check(`${high >> 8}.${high & 255}.${low >> 8}.${low & 255}`, 'ipv4');
	}
	return !IPV6_GLOBAL.check(ip, 'ipv6') || IPV6_NOT_GLOBAL.check(ip, 'ipv6');
}

/** The addresses of a responder host, all of them checked; fails when any is not public. */
async function publicAddresses(hostname: string): Promise<dns.LookupAddress[]> {
	const host = hostname.replace(/^\[|\]$/g, '');
	const family = net.isIP(host);
	const addresses = family
		? [{ address: host, family }]
		: await dns.promises.lookup(host, { all: true });
	if (addresses.length === 0) {
		throw new Error('Responder host did not resolve');
	}
	for (const { address } of addresses) {
		if (isPrivateAddress(address)) {
			throw new Error('Responder host resolves to a private network address');
		}
	}
	return addresses;
}

/**
 * A socket `lookup` answering with `addresses` and never asking DNS again, so a name cannot
 * resolve to one address for the check and to another for the connection.
 */
function pinnedLookup(addresses: dns.LookupAddress[]): net.LookupFunction {
	return (_hostname, options, callback) => {
		const [first] = addresses;
		if (options.all || !first) callback(null, addresses);
		else callback(null, first.address, first.family);
	};
}

function postRequest(
	url: URL,
	body: Buffer,
	contentType: string,
	lookup: net.LookupFunction | undefined,
	signal: AbortSignal,
): Promise<IncomingMessage> {
	return new Promise((resolve, reject) => {
		const request = (url.protocol === 'https:' ? https : http).request(
			url,
			{
				method: 'POST',
				headers: { 'Content-Type': contentType, 'Content-Length': body.byteLength },
				// A pooled socket could lead to an address this request never checked.
				agent: false,
				signal,
				...(lookup ? { lookup } : {}),
			},
			resolve,
		);
		request.on('error', reject);
		request.end(body);
	});
}

async function readCapped(response: IncomingMessage): Promise<Buffer> {
	if (Number(response.headers['content-length'] ?? 0) > OCSP_MAX_RESPONSE_BYTES) {
		response.destroy();
		throw new Error('Response too large');
	}
	const chunks: Buffer[] = [];
	let total = 0;
	for await (const chunk of response as AsyncIterable<Buffer>) {
		total += chunk.byteLength;
		if (total > OCSP_MAX_RESPONSE_BYTES) {
			response.destroy();
			throw new Error('Response too large');
		}
		chunks.push(chunk);
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
	const lookup = options.allowPrivateNetwork
		? undefined
		: pinnedLookup(await publicAddresses(parsed.hostname));
	const signal = AbortSignal.timeout(OCSP_TIMEOUT_MS);
	const response = await postRequest(parsed, body, contentType, lookup, signal);
	const status = response.statusCode ?? 0;
	if (status >= 300 && status < 400) {
		response.destroy();
		throw new Error('Redirects are not followed');
	}
	return readCapped(response);
}
