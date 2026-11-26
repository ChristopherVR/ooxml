import { partNameKey } from '../opc/part-names';
import { relationshipsPartFor } from '../opc/relationships';
import { elements } from '../xml/index';
import { fail } from './package-common';
import type { VisioPackage } from './package';

function extension(path: string): string | undefined {
	return /\.([^./]+)$/.exec(path)?.[1]?.toLowerCase();
}

/** Remove the master and dependencies whose last incoming relationship was removed. */
export async function removeStencilMasterParts({
	pkg,
	parts,
	types,
	master,
	index,
	relId,
	check,
}: {
	pkg: VisioPackage;
	parts: Map<string, Uint8Array>;
	types: Element;
	master: string;
	index: string;
	relId: string | null | undefined;
	check: () => void;
}): Promise<void> {
	const outgoing = new Map<string, string[]>();
	const incoming = new Map<string, number>();
	for (const source of ['', ...parts.keys()]) {
		check();
		if (!parts.has(relationshipsPartFor(source))) continue;
		const targets: string[] = [];
		for (const rel of (await pkg.relationships(source)).values()) {
			if (rel.mode === 'External' || (source === index && rel.id === relId)) continue;
			targets.push(rel.target);
			if (source !== rel.target) incoming.set(rel.target, (incoming.get(rel.target) ?? 0) + 1);
		}
		outgoing.set(source, targets);
	}
	if (incoming.has(master)) fail('UNSUPPORTED_EDIT_PACKAGE', 'Another part uses this master part.');
	const removed = new Set<string>();
	const removedExtensions = new Set<string | undefined>();
	const queue = [master];
	for (const source of queue) {
		check();
		for (const path of [source, relationshipsPartFor(source)]) {
			if (!parts.delete(path)) continue;
			removed.add(partNameKey(path));
			removedExtensions.add(extension(path));
		}
		for (const target of outgoing.get(source) ?? []) {
			if (source === target) continue;
			const count = (incoming.get(target) ?? 0) - 1;
			incoming.set(target, count);
			if (count === 0) queue.push(target);
		}
	}
	const remainingExtensions = new Set([...parts.keys()].map(extension));
	for (const node of elements(types)) {
		const part = node.getAttribute('PartName');
		const ext = node.getAttribute('Extension')?.toLowerCase();
		if (
			(node.localName === 'Override' && part && removed.has(partNameKey(part))) ||
			(node.localName === 'Default' &&
				ext &&
				removedExtensions.has(ext) &&
				!remainingExtensions.has(ext))
		)
			types.removeChild(node);
	}
}
