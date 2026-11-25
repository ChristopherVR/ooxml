/** Per-equation limits, independent of the limits for the enclosing document part. */
export const MAX_LATEX_LENGTH = 16_384;
export const MAX_OMML_LENGTH = 262_144;
export const MAX_OMML_NODES = 16_384;

/** Admit an already-parsed equation before cloning, serializing or converting it. */
export function checkEquationTree(root: Element): void {
	let remainingNodes = MAX_OMML_NODES;
	let remainingLength = MAX_OMML_LENGTH;
	let node: Node = root;
	let depth = 1;
	while (true) {
		if (depth > 512) throw new Error('OMML nesting exceeds 512 levels');
		if (--remainingNodes < 0) throw new Error('OMML exceeds the node limit');
		remainingLength -= node.nodeName.length * 2 + (node.nodeValue?.length ?? 0) + 5;
		if (node.nodeType === 1) {
			for (const { name, value } of Array.from((node as Element).attributes)) {
				remainingLength -= name.length + value.length + 4;
				if (remainingLength < 0) throw new Error('OMML exceeds the size limit');
			}
		}
		if (remainingLength < 0) throw new Error('OMML exceeds the size limit');
		if (node.firstChild) {
			node = node.firstChild;
			depth++;
			continue;
		}
		while (node !== root && !node.nextSibling && node.parentNode) {
			node = node.parentNode;
			depth--;
		}
		if (node === root || !node.nextSibling) return;
		node = node.nextSibling;
	}
}
