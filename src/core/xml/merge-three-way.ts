import { mergeTextThreeWay } from '../text/merge-three-way';
import { buildXml, elements, parseXml } from './xml';

export interface XmlMergeOptions {
	keyElement?: (element: Element, parent: Element) => string | undefined;
	/** Singleton child position; equal positions are mutually exclusive alternatives. */
	childOrder?: (element: Element, parent: Element) => number | undefined;
	allowPositional?: (parent: Element) => boolean;
	allowText?: (element: Element) => boolean;
	resolveText?: (base: Element, local: Element, remote: Element) => string | undefined;
}
export type XmlMergeResult =
	| { status: 'merged'; xml: string }
	| { status: 'unresolved'; issues: readonly string[] };
const tag = (node: Element) => `${node.namespaceURI ?? ''}|${node.localName ?? node.tagName}`;
const xml = (node: Element | undefined) => (node ? buildXml(node) : undefined);

function childrenByKey(
	parent: Element,
	options: XmlMergeOptions,
	repeated: Set<string>,
): Map<string, Element> {
	const result = new Map<string, Element>();
	const counts = new Map<string, number>();
	for (const child of elements(parent)) {
		const name = tag(child),
			explicit = options.keyElement?.(child, parent);
		const index = counts.get(name) ?? 0;
		counts.set(name, index + 1);
		if (!explicit && repeated.has(name) && !options.allowPositional?.(parent))
			throw new Error('Ambiguous repeated elements');
		const key = explicit ? `${name}@${explicit}` : `${name}#${index}`;
		if (result.has(key)) throw new Error('Duplicate element identity');
		result.set(key, child);
	}
	return result;
}

/** Combines peer and schema constraints; only keyed collections admit unconstrained ties. */
function orderKeys(
	base: readonly string[],
	local: readonly string[],
	remote: readonly string[],
	isKeyed: (key: string) => boolean,
	childOrder: (key: string) => number | undefined,
): string[] | undefined {
	const same = (a: readonly string[], b: readonly string[]) =>
		a.length === b.length && a.every((key, index) => key === b[index]);
	if (same(base, local)) return [...remote];
	if (same(base, remote) || same(local, remote)) return [...local];
	// A topological union preserves both peers' independent inserts while rejecting incompatible moves.
	const keys = [...new Set([...remote, ...local])].sort();
	const edges = new Map(keys.map((key) => [key, new Set<string>()]));
	const degree = new Map(keys.map((key) => [key, 0]));
	const addSequence = (sequence: readonly string[]) => {
		for (let i = 1; i < sequence.length; i++) {
			const previous = sequence[i - 1]!,
				next = sequence[i]!;
			if (!edges.get(previous)!.has(next)) {
				edges.get(previous)!.add(next);
				degree.set(next, degree.get(next)! + 1);
			}
		}
	};
	for (const sequence of [local, remote]) addSequence(sequence);
	const ranked = keys.flatMap((key) => {
		const rank = childOrder(key);
		return rank === undefined ? [] : [{ key, rank }];
	});
	if (
		ranked.some(({ rank }) => !Number.isFinite(rank)) ||
		new Set(ranked.map(({ rank }) => rank)).size !== ranked.length
	)
		return undefined;
	addSequence(ranked.sort((a, b) => a.rank - b.rank).map(({ key }) => key));
	const out: string[] = [];
	const ready = keys.filter((key) => degree.get(key) === 0);
	while (ready.length) {
		// A keyed collection admits deterministic insert order. An unkeyed sequence does not.
		if (ready.length > 1 && ready.some((key) => !isKeyed(key))) return undefined;
		const next = ready.sort().shift()!;
		out.push(next);
		for (const key of edges.get(next)!) {
			degree.set(key, degree.get(key)! - 1);
			if (degree.get(key) === 0) ready.push(key);
		}
	}
	return out.length === keys.length ? out : undefined;
}

/** Conservative namespace-aware XML merge. Product adapters explicitly admit positional lists. */
export function mergeXmlThreeWay(
	baseXml: string,
	localXml: string,
	remoteXml: string,
	options: XmlMergeOptions = {},
): XmlMergeResult {
	if (localXml === remoteXml || baseXml === remoteXml) return { status: 'merged', xml: localXml };
	if (baseXml === localXml) return { status: 'merged', xml: remoteXml };
	const issues: string[] = [];
	try {
		const base = parseXml(baseXml, { warningsAsErrors: true });
		const local = parseXml(localXml, { warningsAsErrors: true });
		const remote = parseXml(remoteXml, { warningsAsErrors: true });
		const copy = (node: Element) => remote.importNode(node, true) as Element;
		function merge(
			b: Element | undefined,
			l: Element | undefined,
			r: Element | undefined,
			path: string,
		): Element | undefined {
			const bx = xml(b),
				lx = xml(l),
				rx = xml(r);
			if (lx === rx || bx === rx) return l ? copy(l) : undefined;
			if (bx === lx) return r ? copy(r) : undefined;
			if (!b && l && r && tag(l) === tag(r) && !elements(l).length && !elements(r).length) {
				const resolved = options.resolveText?.(r, l, r);
				if (resolved !== undefined) {
					const node = copy(r);
					node.textContent = resolved;
					return node;
				}
			}
			if (!b || !l || !r || tag(b) !== tag(l) || tag(b) !== tag(r)) {
				issues.push(`${path}: overlapping structural changes`);
				return undefined;
			}
			const out = remote.createElementNS(r.namespaceURI, r.tagName);
			const attrs = (node: Element) =>
				new Map(
					Array.from(node.attributes).map((a) => [
						`${a.namespaceURI ?? ''}|${a.localName ?? a.name}`,
						a,
					]),
				);
			const ba = attrs(b),
				la = attrs(l),
				ra = attrs(r);
			for (const key of new Set([...ba.keys(), ...la.keys(), ...ra.keys()])) {
				const old = ba.get(key)?.value,
					left = la.get(key)?.value,
					right = ra.get(key)?.value;
				const value = left === right || right === old ? left : left === old ? right : undefined;
				if (left !== right && left !== old && right !== old) {
					issues.push(`${path}: overlapping attribute ${key}`);
					continue;
				}
				const source = ra.get(key) ?? la.get(key) ?? ba.get(key)!;
				if (value !== undefined) out.setAttributeNS(source.namespaceURI, source.name, value);
			}
			const lists = [b, l, r].map(elements);
			if (lists.every((list) => list.length === 0)) {
				if (
					[b, l, r].some((node) =>
						Array.from(node.childNodes).some(
							(child) => child.nodeType !== 3 && child.nodeType !== 4,
						),
					)
				) {
					issues.push(`${path}: non-text leaf content requires explicit preservation`);
					return undefined;
				}
				const texts = [b, l, r].map((node) => node.textContent ?? '');
				const override = options.resolveText?.(b, l, r);
				const value =
					override !== undefined
						? { status: 'merged' as const, text: override }
						: texts[1] === texts[2] || texts[0] === texts[2]
							? { status: 'merged' as const, text: texts[1]! }
							: texts[0] === texts[1]
								? { status: 'merged' as const, text: texts[2]! }
								: options.allowText?.(b)
									? mergeTextThreeWay(texts[0]!, texts[1]!, texts[2]!)
									: { status: 'unresolved' as const };
				if (value.status !== 'merged') issues.push(`${path}: overlapping text changes`);
				else if (value.text) out.appendChild(remote.createTextNode(value.text));
				return out;
			}
			if (
				[b, l, r].some((node) =>
					Array.from(node.childNodes).some(
						(child) =>
							child.nodeType !== 1 && (child.nodeType !== 3 || Boolean(child.textContent?.trim())),
					),
				)
			) {
				issues.push(`${path}: mixed content cannot be matched safely`);
				return undefined;
			}
			const repeated = new Set<string>();
			const groups = lists.map((list) => {
				const grouped = new Map<string, Element[]>();
				for (const child of list) {
					const name = tag(child),
						group = grouped.get(name) ?? [];
					group.push(child);
					grouped.set(name, group);
					if (group.length > 1) repeated.add(name);
				}
				return grouped;
			});
			for (const name of repeated) {
				const rows = groups.map((group) => group.get(name) ?? []);
				if (rows.some((row) => row.some((child) => options.keyElement?.(child, b) !== undefined)))
					continue;
				if (rows.some((row) => row.length !== rows[0]!.length)) {
					issues.push(`${path}: unkeyed structural changes`);
					return undefined;
				}
				const before = rows[0]!.map((child) => buildXml(child));
				const known = new Set(before);
				for (const row of rows.slice(1))
					if (
						row.some((child, i) => {
							const value = buildXml(child);
							return before[i] !== value && known.has(value);
						})
					) {
						issues.push(`${path}: reordered unkeyed elements`);
						return undefined;
					}
			}
			let maps: Map<string, Element>[];
			try {
				maps = [b, l, r].map((parent) => childrenByKey(parent, options, repeated));
			} catch {
				issues.push(`${path}: ambiguous element identities`);
				return undefined;
			}
			const [bm, lm, rm] = maps as [
				Map<string, Element>,
				Map<string, Element>,
				Map<string, Element>,
			];
			const merged = new Map<string, Element>();
			for (const key of new Set([...bm.keys(), ...lm.keys(), ...rm.keys()])) {
				const node = merge(bm.get(key), lm.get(key), rm.get(key), `${path}/${key}`);
				if (node) merged.set(key, node);
			}
			const surviving = (map: Map<string, Element>) =>
				[...map.keys()].filter((key) => merged.has(key));
			const order = orderKeys(
				surviving(bm),
				surviving(lm),
				surviving(rm),
				(key) => Boolean(options.keyElement?.(merged.get(key)!, out)),
				(key) => options.childOrder?.(merged.get(key)!, out),
			);
			if (!order) {
				issues.push(`${path}: incompatible element order`);
				return undefined;
			}
			for (const key of order) out.appendChild(merged.get(key)!);
			return out;
		}
		const root = merge(base.documentElement, local.documentElement, remote.documentElement, 'root');
		if (!root || issues.length) return { status: 'unresolved', issues };
		remote.replaceChild(root, remote.documentElement);
		return { status: 'merged', xml: buildXml(remote) };
	} catch {
		return { status: 'unresolved', issues: ['Invalid or unsupported XML input'] };
	}
}
