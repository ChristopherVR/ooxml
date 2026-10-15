import dns from 'node:dns';

import { afterEach, describe, expect, it, vi } from 'vitest';

import { OCSP_MAX_RESPONSE_BYTES, httpPost, isPrivateAddress } from './ocsp-network';

function mockLookup(...addresses: string[]) {
	return vi
		.spyOn(dns.promises, 'lookup')
		.mockResolvedValue(
			addresses.map((address) => ({ address, family: address.includes(':') ? 6 : 4 })) as never,
		);
}

afterEach(() => {
	vi.restoreAllMocks();
	vi.unstubAllGlobals();
});

describe('isPrivateAddress', () => {
	it.each([
		'127.0.0.1',
		'10.1.2.3',
		'172.16.0.1',
		'192.168.1.1',
		'169.254.169.254',
		'0.0.0.0',
		'::1',
		'::',
		'fe80::1',
		'fd00::1',
		'::ffff:127.0.0.1',
		'::ffff:7f00:1',
	])('flags %s', (ip) => expect(isPrivateAddress(ip)).toBe(true));

	it.each(['8.8.8.8', '93.184.216.34', '2606:4700::1111'])('allows %s', (ip) =>
		expect(isPrivateAddress(ip)).toBe(false),
	);
});

describe('httpPost', () => {
	const body = Buffer.from([1, 2, 3]);

	it('rejects non-http protocols', async () => {
		await expect(httpPost('file:///etc/passwd', body, 't')).rejects.toThrow('Unsupported');
	});

	it('rejects hosts resolving to private addresses without fetching', async () => {
		mockLookup('93.184.216.34', '169.254.169.254');
		const fetchMock = vi.fn();
		vi.stubGlobal('fetch', fetchMock);
		await expect(httpPost('http://ocsp.example/', body, 't')).rejects.toThrow('private');
		expect(fetchMock).not.toHaveBeenCalled();
	});

	it('rejects literal private IPs', async () => {
		vi.stubGlobal('fetch', vi.fn());
		await expect(httpPost('http://127.0.0.1:8080/', body, 't')).rejects.toThrow('private');
		await expect(httpPost('http://[::1]/', body, 't')).rejects.toThrow('private');
	});

	it('allows private hosts with allowPrivateNetwork', async () => {
		const fetchMock = vi.fn().mockResolvedValue(new Response(new Uint8Array([9])));
		vi.stubGlobal('fetch', fetchMock);
		const out = await httpPost('http://127.0.0.1/', body, 't', { allowPrivateNetwork: true });
		expect([...out]).toEqual([9]);
	});

	it('posts to public hosts without following redirects', async () => {
		mockLookup('93.184.216.34');
		const fetchMock = vi.fn().mockResolvedValue(new Response(new Uint8Array([7])));
		vi.stubGlobal('fetch', fetchMock);
		await httpPost('https://ocsp.example/x', body, 'application/ocsp-request');
		expect(fetchMock.mock.calls[0]?.[1]).toMatchObject({ method: 'POST', redirect: 'manual' });
	});

	it('refuses redirect responses', async () => {
		mockLookup('93.184.216.34');
		vi.stubGlobal(
			'fetch',
			vi
				.fn()
				.mockResolvedValue(new Response(null, { status: 302, headers: { location: 'http://x' } })),
		);
		await expect(httpPost('http://ocsp.example/', body, 't')).rejects.toThrow('Redirects');
	});

	it('aborts responses over 1 MiB', async () => {
		mockLookup('93.184.216.34');
		const big = new Uint8Array(OCSP_MAX_RESPONSE_BYTES + 1);
		vi.stubGlobal('fetch', vi.fn().mockResolvedValue(new Response(big)));
		await expect(httpPost('http://ocsp.example/', body, 't')).rejects.toThrow('too large');
	});
});
