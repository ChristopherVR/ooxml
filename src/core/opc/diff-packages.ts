import JSZip from 'jszip';

/** The parts that differ between two OOXML packages, by part name. */
export interface OpcPackageDiff {
	/** In `other` only. */
	readonly added: readonly string[];
	/** In `base` only. */
	readonly removed: readonly string[];
	/** In both, with different uncompressed content. */
	readonly changed: readonly string[];
}

const sameBytes = (a: Uint8Array, b: Uint8Array) =>
	a.length === b.length && a.every((value, index) => value === b[index]);

/**
 * Compare two packages part by part on their uncompressed content, so re-compression or a
 * different part order is not reported. For diagnostics (what changed in a cloud copy since it was
 * opened); it reads every part, so keep it off hot paths.
 */
export async function diffOpcPackages(
	base: Uint8Array,
	other: Uint8Array,
): Promise<OpcPackageDiff> {
	const [left, right] = await Promise.all([JSZip.loadAsync(base), JSZip.loadAsync(other)]);
	const names = (zip: JSZip) =>
		new Set(
			Object.values(zip.files)
				.filter((entry) => !entry.dir)
				.map((entry) => entry.name),
		);
	const leftNames = names(left);
	const rightNames = names(right);
	const changed: string[] = [];
	for (const name of leftNames) {
		if (!rightNames.has(name)) continue;
		const [a, b] = await Promise.all([
			left.file(name)!.async('uint8array'),
			right.file(name)!.async('uint8array'),
		]);
		if (!sameBytes(a, b)) changed.push(name);
	}
	return {
		added: [...rightNames].filter((name) => !leftNames.has(name)).sort(),
		removed: [...leftNames].filter((name) => !rightNames.has(name)).sort(),
		changed: changed.sort(),
	};
}
