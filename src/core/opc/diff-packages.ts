import type JSZip from 'jszip';
import { openZip, type ZipLimits } from './zip-limits';

/** The parts that differ between two OOXML packages, by part name. */
export interface OpcPackageDiff {
	/** In `other` only. */
	readonly added: readonly string[];
	/** In `base` only. */
	readonly removed: readonly string[];
	/** In both, with different uncompressed content. */
	readonly changed: readonly string[];
}

/** Options of {@link diffOpcPackages}: the package size limits, plus a part comparison override. */
export interface DiffOpcPackagesOptions extends ZipLimits {
	/** Whether two different byte sequences of `part` hold the same content. */
	equivalent?(part: string, a: Uint8Array, b: Uint8Array): boolean;
}

const sameBytes = (a: Uint8Array, b: Uint8Array) =>
	a.length === b.length && a.every((value, index) => value === b[index]);

/**
 * Compare two packages part by part on their uncompressed content, so re-compression or a
 * different part order is not reported. For diagnostics (what changed in a cloud copy since it was
 * opened); it reads every part, so keep it off hot paths. Rejects a package over the
 * {@link ZipLimits} in `options`, by default those of {@link openZip}; `equivalent` can accept two
 * different byte sequences of one part as the same content.
 */
export async function diffOpcPackages(
	base: Uint8Array,
	other: Uint8Array,
	options: DiffOpcPackagesOptions = {},
): Promise<OpcPackageDiff> {
	const { equivalent, ...limits } = options;
	const [left, right] = await Promise.all([
		openZip(base, { label: 'Package', ...limits }),
		openZip(other, { label: 'Package', ...limits }),
	]);
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
		if (!sameBytes(a, b) && !equivalent?.(name, a, b)) changed.push(name);
	}
	return {
		added: [...rightNames].filter((name) => !leftNames.has(name)).sort(),
		removed: [...leftNames].filter((name) => !rightNames.has(name)).sort(),
		changed: changed.sort(),
	};
}
