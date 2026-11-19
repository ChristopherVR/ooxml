import dns from 'node:dns';
import http from 'node:http';
import net, { type AddressInfo } from 'node:net';

import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';

import { OCSP_MAX_RESPONSE_BYTES, httpPost, isPrivateAddress } from './ocsp-network';

const answer = (...addresses: string[]) =>
	addresses.map((address) => ({ address, family: address.includes(':') ? 6 : 4 }));

function mockLookup(...addresses: string[]) {
	return vi.spyOn(dns.promises, 'lookup').mockResolvedValue(answer(...addresses) as never);
}

let server: http.Server;
let received: { method?: string | undefined; type?: string | undefined; body: Buffer }[];
let respond: http.RequestListener;

beforeEach(async () => {
	received = [];
	respond = (_request, response) => response.end(Buffer.from([7]));
	server = http.createServer((request, response) => {
		const chunks: Buffer[] = [];
		request.on('data', (chunk: Buffer) => chunks.push(chunk));
		request.on('end', () => {
			received.push({
				method: request.method,
				type: request.headers['content-type'],
				body: Buffer.concat(chunks),
			});
			respond(request, response);
		});
	});
	await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve));
});

afterEach(async () => {
	vi.restoreAllMocks();
	server.closeAllConnections();
	await new Promise((resolve) => server.close(resolve));
});

const port = () => (server.address() as AddressInfo).port;

interface ConnectOptions {
	lookup?: net.LookupFunction;
}

/**
 * Sends the guarded request's socket to the local test server, whatever public address it was
 * pinned to, and keeps the options it was opened with.
 */
function captureConnection() {
	const seen: ConnectOptions[] = [];
	const agent = http.Agent.prototype as unknown as {
		createConnection(options: ConnectOptions, callback: () => void): net.Socket;
	};
	vi.spyOn(agent, 'createConnection').mockImplementation((options, callback) => {
		seen.push(options);
		return net.createConnection({ host: '127.0.0.1', port: port() }, callback);
	});
	return seen;
}

function lookupAll(lookup: net.LookupFunction | undefined, hostname: string) {
	return new Promise<unknown>((resolve, reject) => {
		lookup?.(hostname, { all: true }, (error, addresses) =>
			error ? reject(error) : resolve(addresses),
		);
	});
}

describe('isPrivateAddress', () => {
	it.each([
		'127.0.0.1',
		'10.1.2.3',
		'172.16.0.1',
		'192.168.1.1',
		'169.254.169.254',
		'100.64.0.1',
		'0.0.0.0',
		'192.0.0.192',
		'192.0.2.1',
		'192.88.99.1',
		'198.18.0.1',
		'198.51.100.1',
		'203.0.113.1',
		'224.0.0.1',
		'240.0.0.1',
		'255.255.255.255',
		'::1',
		'::',
		'fe80::1',
		'fe80::1%eth0',
		'fd00::1',
		'ff02::1',
		'::7f00:1',
		'::ffff:127.0.0.1',
		'::ffff:7f00:1',
		'0:0:0:0:0:ffff:7f00:1',
		'64:ff9b::7f00:1',
		'64:ff9b::127.0.0.1',
		'64:ff9b:1::808:808',
		'2001::1',
		'2001:db8::1',
		'2002:7f00:1::',
		'2002:808:808::',
		'3fff::1',
		'not-an-address',
	])('refuses %s', (ip) => expect(isPrivateAddress(ip)).toBe(true));

	it.each([
		'8.8.8.8',
		'93.184.216.34',
		'[2606:4700::1111]',
		'2606:4700::1111',
		'::ffff:8.8.8.8',
		'::ffff:808:808',
		'64:ff9b::808:808',
	])('allows %s', (ip) => expect(isPrivateAddress(ip)).toBe(false));
});

describe('httpPost', () => {
	const body = Buffer.from([1, 2, 3]);

	it('rejects non-http protocols', async () => {
		await expect(httpPost('file:///etc/passwd', body, 't')).rejects.toThrow('Unsupported');
	});

	it('connects to no host when any of its addresses is not public', async () => {
		mockLookup('93.184.216.34', '169.254.169.254');
		const connections = captureConnection();
		await expect(httpPost('http://ocsp.example/', body, 't')).rejects.toThrow('private');
		expect(connections).toEqual([]);
	});

	it('rejects literal private addresses', async () => {
		const connections = captureConnection();
		await expect(httpPost(`http://127.0.0.1:${port()}/`, body, 't')).rejects.toThrow('private');
		await expect(httpPost('http://[::1]/', body, 't')).rejects.toThrow('private');
		await expect(httpPost('http://[64:ff9b::7f00:1]/', body, 't')).rejects.toThrow('private');
		expect(connections).toEqual([]);
		expect(received).toEqual([]);
	});

	it('allows private hosts with allowPrivateNetwork', async () => {
		const url = `http://127.0.0.1:${port()}/`;
		const out = await httpPost(url, body, 't', { allowPrivateNetwork: true });
		expect([...out]).toEqual([7]);
	});

	it('connects to the checked address when the name later resolves elsewhere', async () => {
		const lookup = vi
			.spyOn(dns.promises, 'lookup')
			.mockResolvedValueOnce(answer('93.184.216.34') as never)
			.mockResolvedValue(answer('127.0.0.1') as never);
		const connections = captureConnection();
		const out = await httpPost('http://ocsp.example/x', body, 'application/ocsp-request');
		expect([...out]).toEqual([7]);
		expect(lookup).toHaveBeenCalledTimes(1);
		expect(await lookupAll(connections[0]?.lookup, 'ocsp.example')).toEqual(
			answer('93.184.216.34'),
		);
	});

	it('sends the request body and content type', async () => {
		mockLookup('93.184.216.34');
		captureConnection();
		await httpPost('http://ocsp.example/x', body, 'application/ocsp-request');
		expect(received).toEqual([{ method: 'POST', type: 'application/ocsp-request', body }]);
	});

	it('refuses redirect responses', async () => {
		mockLookup('93.184.216.34');
		captureConnection();
		respond = (_request, response) => {
			response.writeHead(302, { location: 'http://127.0.0.1/' }).end();
		};
		await expect(httpPost('http://ocsp.example/', body, 't')).rejects.toThrow('Redirects');
		expect(received).toHaveLength(1);
	});

	it('aborts a response declared over 1 MiB', async () => {
		mockLookup('93.184.216.34');
		captureConnection();
		respond = (_request, response) => {
			response.writeHead(200, { 'content-length': OCSP_MAX_RESPONSE_BYTES + 1 });
			response.flushHeaders();
		};
		await expect(httpPost('http://ocsp.example/', body, 't')).rejects.toThrow('too large');
	});

	it('aborts a streamed response that grows past 1 MiB', async () => {
		mockLookup('93.184.216.34');
		captureConnection();
		respond = (_request, response) => {
			response.write(Buffer.alloc(OCSP_MAX_RESPONSE_BYTES));
			response.write(Buffer.alloc(1024));
		};
		await expect(httpPost('http://ocsp.example/', body, 't')).rejects.toThrow('too large');
	});
});
