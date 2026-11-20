// Runs one corpus case in a child process whose heap the parent capped: a heap a case exhausts
// ends this process, never the test runner. Every way out to the network is replaced first, so an
// attempt is recorded and fails instead of leaving the machine.
import dns from 'node:dns';
import http from 'node:http';
import https from 'node:https';
import { syncBuiltinESMExports } from 'node:module';
import net from 'node:net';
import { getHeapStatistics } from 'node:v8';

const attempts = [];
const refuse = (what) => {
	attempts.push(what);
	throw new Error(`network access refused by the security corpus: ${what}`);
};
globalThis.fetch = async (input) => refuse(`fetch ${String(input?.url ?? input)}`);
globalThis.WebSocket = class {
	constructor(url) {
		refuse(`WebSocket ${String(url)}`);
	}
};
globalThis.EventSource = class {
	constructor(url) {
		refuse(`EventSource ${String(url)}`);
	}
};
net.Socket.prototype.connect = function () {
	refuse('socket connect');
};
for (const module of [http, https]) {
	module.request = (target) => refuse(`${module === http ? 'http' : 'https'} ${String(target)}`);
	module.get = module.request;
}
for (const name of ['lookup', 'resolve', 'resolve4', 'resolve6', 'resolveAny']) {
	dns[name] = (host) => refuse(`dns ${name} ${String(host)}`);
	dns.promises[name] = async (host) => refuse(`dns ${name} ${String(host)}`);
}
syncBuiltinESMExports();

const [bundle, index] = process.argv.slice(2);
const { cases } = await import(bundle);
const entry = cases[Number(index)];
process.send({ type: 'ready' });
const start = performance.now();
let value;
let error;
try {
	value = await entry.run();
} catch (cause) {
	error = cause instanceof Error ? (cause.stack ?? cause.message) : String(cause);
}
const ms = performance.now() - start;
const heapMb = getHeapStatistics().used_heap_size / 2 ** 20;
process.send({ type: 'done', value, error, ms, heapMb, network: attempts }, () => process.exit(0));
