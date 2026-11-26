import { xpathSelector, type XPathSelector } from './xpath-selector';

let selector: XPathSelector | undefined;
let pending: Promise<void> | undefined;
let unavailable = false;
const consumers = new WeakSet<object>();

/** Node and Bun keep the synchronous calculation API without loading XPath at import time. */
export function loadedXPath(): XPathSelector | undefined {
	if (selector || unavailable) return selector;
	const module = globalThis.process?.getBuiltinModule?.('module');
	if (module) {
		const load = module.createRequire(
			typeof __filename === 'string' ? __filename : import.meta.url,
		);
		selector = xpathSelector(load('xpath'));
	}
	return selector;
}

/** Bundlers share this module between workbook loading, editing and the formula engine. */
export function prepareXPath(consumer?: object): Promise<void> | undefined {
	if (loadedXPath() || unavailable) return undefined;
	if (consumer) consumers.add(consumer);
	return (pending ??= import('xpath').then(
		(api) => {
			selector = xpathSelector(api);
			pending = undefined;
		},
		() => {
			unavailable = true;
			pending = undefined;
		},
	));
}

/** An in-flight load, without starting one. */
export const pendingXPath = (consumer: object): Promise<void> | undefined =>
	consumers.has(consumer) ? pending : undefined;
