// Heap snapshot measurements for the browser leak tests: which detached DOM nodes are still alive
// after garbage collection, and how many bytes they keep alive.

/**
 * Retained-set report of a parsed V8 heap snapshot for the detached DOM nodes, or only those whose
 * snapshot name (`<tag attributes...>`) matches `name`. `retainedBytes` counts every object that
 * is alive only because of the selected nodes (their exact retained size as a set).
 * @param {any} snapshot
 * @param {RegExp} [name]
 */
export function retainedBy(snapshot, name) {
	const meta = snapshot.snapshot.meta;
	const nodeFields = meta.node_fields;
	const edgeFields = meta.edge_fields;
	const NF = nodeFields.length;
	const EF = edgeFields.length;
	const nType = nodeFields.indexOf('type');
	const nName = nodeFields.indexOf('name');
	const nSize = nodeFields.indexOf('self_size');
	const nEdges = nodeFields.indexOf('edge_count');
	const nDetached = nodeFields.indexOf('detachedness');
	if (nDetached < 0)
		throw new Error('This heap snapshot does not record which DOM nodes are detached');
	const eType = edgeFields.indexOf('type');
	const eTo = edgeFields.indexOf('to_node');
	const weak = meta.edge_types[0].indexOf('weak');
	const textual = new Set(
		['string', 'concatenated string', 'sliced string']
			.map((type) => meta.node_types[0].indexOf(type))
			.filter((index) => index >= 0),
	);
	const { nodes, edges, strings } = snapshot;
	const count = nodes.length / NF;
	const firstEdge = new Uint32Array(count + 1);
	for (let node = 0, edge = 0; node < count; node++) {
		firstEdge[node] = edge;
		edge += nodes[node * NF + nEdges];
		firstEdge[node + 1] = edge;
	}
	/** Marks what is reachable over strong edges from `starts`, never entering `blocked`. */
	const walk = (starts, blocked) => {
		const seen = new Uint8Array(count);
		const queue = new Uint32Array(count);
		let head = 0;
		let tail = 0;
		for (const start of starts) {
			seen[start] = 1;
			queue[tail++] = start;
		}
		while (head < tail) {
			const node = queue[head++];
			for (let edge = firstEdge[node]; edge < firstEdge[node + 1]; edge++) {
				if (edges[edge * EF + eType] === weak) continue;
				const to = edges[edge * EF + eTo] / NF;
				if (seen[to] || blocked[to]) continue;
				seen[to] = 1;
				queue[tail++] = to;
			}
		}
		return seen;
	};
	// Node 0 is the snapshot's synthetic root; a node reached only through weak edges is garbage.
	const alive = walk([0], new Uint8Array(count));
	const matched = new Uint8Array(count);
	let matches = 0;
	for (let node = 0; node < count; node++) {
		if (textual.has(nodes[node * NF + nType])) continue;
		// Detachedness: 0 not a DOM node, 1 attached, 2 detached.
		if (!alive[node] || nodes[node * NF + nDetached] !== 2) continue;
		if (!name || name.test(strings[nodes[node * NF + nName]])) {
			matched[node] = 1;
			matches++;
		}
	}
	const withoutMatched = walk([0], matched);
	const fromMatched = walk(
		[...matched.keys()].filter((node) => matched[node]),
		withoutMatched,
	);
	let retainedBytes = 0;
	for (let node = 0; node < count; node++) {
		if (fromMatched[node]) retainedBytes += nodes[node * NF + nSize];
	}
	return { count: matches, retainedBytes };
}

/**
 * Runs in the page (`page.evaluate(openAndCloseEditor, options)`): creates a `tag` element marked
 * `data-leak-probe`, loads `bytes` into it, waits `settle` ms for it to paint, removes it and
 * returns how many of its painted elements (shadow trees included) reference a `url(#id)` paint
 * server, clip path, mask, filter or marker while having element children of their own.
 * @param {{ tag: string, bytes: number[], fileName: string, settle?: number }} options
 */
export async function openAndCloseEditor({ tag, bytes, fileName, settle = 1500 }) {
	const element = document.createElement(tag);
	element.setAttribute('data-leak-probe', '');
	element.style.cssText = 'position:fixed;inset:0;z-index:9999;display:block';
	document.body.append(element);
	await element.load(new Uint8Array(bytes), fileName);
	await new Promise((resolve) => setTimeout(resolve, settle));
	const REFERENCE = /url\(\s*['"]?#/;
	const ATTRIBUTES = [
		'fill',
		'stroke',
		'clip-path',
		'mask',
		'filter',
		'marker-start',
		'marker-mid',
		'marker-end',
		'style',
	];
	let references = 0;
	const visit = (root) => {
		for (const node of root.querySelectorAll('*')) {
			if (
				node.firstElementChild &&
				ATTRIBUTES.some((name) => REFERENCE.test(node.getAttribute(name) ?? ''))
			)
				references++;
			if (node.shadowRoot) visit(node.shadowRoot);
		}
	};
	visit(element.shadowRoot ?? element);
	element.remove();
	return references;
}

/**
 * Collects garbage in the page behind `session` (a Playwright CDP session) and takes a heap
 * snapshot. Reports how many detached DOM nodes matching `pattern` are alive, the bytes kept
 * alive only by detached DOM, and the JS heap in use.
 * @param {import('./heap-snapshot.d.mts').HeapSession} session
 * @param {RegExp} pattern
 */
export async function heapReport(session, pattern) {
	await session.send('HeapProfiler.enable');
	for (let i = 0; i < 4; i++) await session.send('HeapProfiler.collectGarbage');
	const { usedSize } = await session.send('Runtime.getHeapUsage');
	const chunks = [];
	const onChunk = (event) => chunks.push(event.chunk);
	session.on('HeapProfiler.addHeapSnapshotChunk', onChunk);
	try {
		await session.send('HeapProfiler.takeHeapSnapshot', { reportProgress: false });
	} finally {
		session.off('HeapProfiler.addHeapSnapshotChunk', onChunk);
	}
	const snapshot = JSON.parse(chunks.join(''));
	return {
		count: retainedBy(snapshot, pattern).count,
		detachedBytes: retainedBy(snapshot).retainedBytes,
		heapBytes: usedSize,
	};
}
