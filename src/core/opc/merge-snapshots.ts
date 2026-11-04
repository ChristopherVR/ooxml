import { mergeXmlThreeWay } from '../xml/merge-three-way';
import { openZip } from './zip-limits';
import { parseRelationships, resolvePartPath } from './relationships';

export type SnapshotMergeResult =
	| { status: 'merged'; bytes: Uint8Array }
	| { status: 'unresolved'; issues: readonly string[] };
export interface PartMergeInput {
	path: string;
	base?: Uint8Array;
	local?: Uint8Array;
	remote?: Uint8Array;
}
export type PartMergeResult =
	| { status: 'merged'; bytes: Uint8Array | undefined }
	| { status: 'unresolved'; issues: readonly string[] };
export interface OpcMergeOptions {
	mergePart?: (input: PartMergeInput) => PartMergeResult | Promise<PartMergeResult>;
}
const equal = (a: Uint8Array | undefined, b: Uint8Array | undefined) =>
	a === b ||
	Boolean(a && b && a.length === b.length && a.every((value, index) => value === b[index]));
const decoder = new TextDecoder('utf-8', { fatal: true });
const encoder = new TextEncoder();

function mergeManifest(input: PartMergeInput): PartMergeResult | undefined {
	const { path, base, local, remote } = input;
	if (!base || !local || !remote) return undefined;
	if (!path.endsWith('.rels') && path !== '[Content_Types].xml' && path !== 'docProps/core.xml')
		return undefined;
	try {
		const result = mergeXmlThreeWay(
			decoder.decode(base),
			decoder.decode(local),
			decoder.decode(remote),
			{
				keyElement: (element) => {
					const attribute =
						element.localName === 'Relationship'
							? 'Id'
							: element.localName === 'Override'
								? 'PartName'
								: element.localName === 'Default'
									? 'Extension'
									: undefined;
					return attribute ? (element.getAttribute(attribute) ?? undefined) : undefined;
				},
				resolveText: (_b, _l, r) =>
					path === 'docProps/core.xml' &&
					((r.namespaceURI === 'http://purl.org/dc/terms/' && r.localName === 'modified') ||
						(r.namespaceURI ===
							'http://schemas.openxmlformats.org/package/2006/metadata/core-properties' &&
							r.localName === 'lastModifiedBy'))
						? (r.textContent ?? '')
						: undefined,
			},
		);
		return result.status === 'merged'
			? { status: 'merged', bytes: encoder.encode(result.xml) }
			: result;
	} catch {
		return { status: 'unresolved', issues: [`${path}: invalid metadata XML`] };
	}
}

/** Three-way OPC part preservation. A product must explicitly resolve any double-edited content. */
export async function mergeOpcSnapshots(
	baseBytes: Uint8Array,
	localBytes: Uint8Array,
	remoteBytes: Uint8Array,
	options: OpcMergeOptions = {},
): Promise<SnapshotMergeResult> {
	if (equal(localBytes, remoteBytes) || equal(baseBytes, remoteBytes))
		return { status: 'merged', bytes: localBytes.slice() };
	if (equal(baseBytes, localBytes)) return { status: 'merged', bytes: remoteBytes.slice() };
	try {
		const [base, local, remote] = await Promise.all([
			openZip(baseBytes, { label: 'Merge' }),
			openZip(localBytes, { label: 'Merge' }),
			openZip(remoteBytes, { label: 'Merge' }),
		]);
		const paths = [
			...new Set(
				[base, local, remote].flatMap((zip) =>
					Object.keys(zip.files).filter((path) => !zip.files[path]!.dir),
				),
			),
		].sort();
		if (paths.some((path) => path.startsWith('_xmlsignatures/')))
			return {
				status: 'unresolved',
				issues: ['Signed packages require explicit signature handling before merging'],
			};
		const issues: string[] = [];
		for (const path of paths) {
			const [b, l, r] = await Promise.all(
				[base, local, remote].map((zip) => zip.file(path)?.async('uint8array')),
			);
			let chosen: Uint8Array | undefined;
			if (equal(l, r) || equal(b, r)) chosen = l;
			else if (equal(b, l)) chosen = r;
			else {
				const input: PartMergeInput = {
					path,
					...(b ? { base: b } : {}),
					...(l ? { local: l } : {}),
					...(r ? { remote: r } : {}),
				};
				const result = mergeManifest(input) ?? (await options.mergePart?.(input));
				if (!result || result.status === 'unresolved') {
					issues.push(...(result?.issues ?? [`${path}: both versions changed the same part`]));
					continue;
				}
				chosen = result.bytes;
			}
			if (chosen) remote.file(path, chosen);
			else remote.remove(path);
		}
		if (issues.length) return { status: 'unresolved', issues };
		for (const path of Object.keys(remote.files).filter((name) => name.endsWith('.rels'))) {
			const source =
				path === '_rels/.rels' ? '' : path.replace(/(^|\/)_rels\/([^/]+)\.rels$/u, '$1$2');
			for (const relation of parseRelationships(
				await remote.file(path)!.async('string'),
			).values()) {
				if (relation.mode === 'External') continue;
				const target = resolvePartPath(source, relation.target.split('#')[0] ?? '');
				if (!remote.file(target) && !remote.file(decodeURIComponent(target)))
					issues.push(`${path}: merged relationship target is missing`);
			}
		}
		if (issues.length) return { status: 'unresolved', issues };
		return {
			status: 'merged',
			bytes: await remote.generateAsync({ type: 'uint8array', compression: 'DEFLATE' }),
		};
	} catch {
		return { status: 'unresolved', issues: ['Invalid or oversized Office package'] };
	}
}
