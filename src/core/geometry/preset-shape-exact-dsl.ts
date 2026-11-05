/**
 * Compact authoring helpers for spec-exact preset geometry.
 *
 * ECMA-376 writes a preset as an `avLst`, a `gdLst` of one-line formulas and a
 * `pathLst` of commands. These helpers let a definition be authored in nearly
 * that shape (`'x1 +- hc 0 dx1'`, `'M x1 y1 L x2 y1 Z'`) so it can be compared
 * line by line with the specification, then expand to the table's types.
 */

import type {
	PresetPath,
	PresetPathCommand,
	PresetShapeGeometryDefinition,
} from './preset-shape-definitions-table';

/** Split `'name op a b c'` into a `gdLst` entry. */
export function g(line: string): { name: string; formula: string; args: string[] } {
	const [name, ...rest] = line.trim().split(/\s+/);
	return { name: name!, formula: rest.join(' '), args: rest.slice(1) };
}

const ARITY: Record<string, number> = { M: 2, L: 2, A: 4, Q: 4, C: 6, Z: 0 };

/**
 * Parse a path mini-language: `M x y`, `L x y`, `A wR hR stAng swAng`,
 * `Q x1 y1 x2 y2`, `C x1 y1 x2 y2 x3 y3` and `Z`, whitespace separated.
 */
export function cmds(src: string): PresetPathCommand[] {
	const tokens = src.trim().split(/\s+/);
	const out: PresetPathCommand[] = [];
	let i = 0;
	while (i < tokens.length) {
		const op = tokens[i++]!;
		const n = ARITY[op];
		if (n === undefined) {
			throw new Error(`Unknown path op ${op}`);
		}
		const a = tokens.slice(i, i + n);
		i += n;
		switch (op) {
			case 'M':
				out.push({ kind: 'moveTo', x: a[0]!, y: a[1]! });
				break;
			case 'L':
				out.push({ kind: 'lnTo', x: a[0]!, y: a[1]! });
				break;
			case 'A':
				out.push({ kind: 'arcTo', wR: a[0]!, hR: a[1]!, stAng: a[2]!, swAng: a[3]! });
				break;
			case 'Q':
				out.push({ kind: 'quadBezTo', x1: a[0]!, y1: a[1]!, x2: a[2]!, y2: a[3]! });
				break;
			case 'C':
				out.push({
					kind: 'cubicBezTo',
					x1: a[0]!,
					y1: a[1]!,
					x2: a[2]!,
					y2: a[3]!,
					x3: a[4]!,
					y3: a[5]!,
				});
				break;
			default:
				out.push({ kind: 'close' });
		}
	}
	return out;
}

/** One `<a:path>`: commands plus the optional `fill` / `stroke` flags. */
export function p(
	src: string,
	flags: { fill?: string; stroke?: boolean; w?: number; h?: number } = {},
): PresetPath {
	return { ...flags, commands: cmds(src) };
}

/** Assemble a definition from the compact pieces. */
export function exactDef(
	name: string,
	avLst: Record<string, number> | undefined,
	gdLines: string[],
	rect: string,
	paths: PresetPath[],
): PresetShapeGeometryDefinition {
	const [l, t, r, b] = rect.trim().split(/\s+/);
	return {
		name,
		...(avLst ? { avLst } : {}),
		gdLst: gdLines.map(g),
		rect: { l: l!, t: t!, r: r!, b: b! },
		pathLst: paths,
	};
}
