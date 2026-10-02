import {
	VisioForeignVectorError,
	limit,
	unsafe,
	type VisioForeignVector,
	type VisioForeignVectorClip,
	type VisioForeignVectorClipPath,
	type VisioForeignVectorLimits,
	type VisioForeignVectorNode,
	type VisioForeignVectorPath,
} from './foreign-vector-types.js';
import {
	finite,
	limitsFor,
	matrix,
	numberList,
	paint,
	record,
	rule,
	type RecordValue,
} from './foreign-vector-values.js';
import { pathCommands } from './foreign-vector-path.js';
import { checkVectorGraph } from './foreign-vector-graph.js';
export { VISIO_FOREIGN_VECTOR_LIMITS, VisioForeignVectorError } from './foreign-vector-types.js';
export type {
	VisioForeignVector,
	VisioForeignVectorClip,
	VisioForeignVectorClipPath,
	VisioForeignVectorCommand,
	VisioForeignVectorGroup,
	VisioForeignVectorLimits,
	VisioForeignVectorMatrix,
	VisioForeignVectorNode,
	VisioForeignVectorPaint,
	VisioForeignVectorPath,
	VisioForeignVectorRule,
} from './foreign-vector-types.js';

const BASE = ['transform', 'clip-path'];
const PAINT = [
	'fill',
	'stroke',
	'fill-rule',
	'stroke-width',
	'stroke-miterlimit',
	'stroke-linecap',
	'stroke-linejoin',
	'opacity',
	'fill-opacity',
	'stroke-opacity',
];
const ATTRS: Readonly<Record<string, readonly string[]>> = {
	svg: ['xmlns', 'width', 'height', 'viewBox', 'style'],
	defs: [],
	g: BASE,
	path: ['d', ...BASE, ...PAINT],
	clipPath: ['id', 'clipPathUnits', 'clip-rule', ...BASE],
};
const ID = /^[A-Za-z_][A-Za-z0-9_.-]{0,127}$/;
type ClipTarget = { clipIndex?: number };
/**
 * Converts a narrowly supported converter-owned SVG tree into a copied, deeply frozen scene.
 * This is an output boundary, not permission to run a converter or evidence of EMF fidelity.
 * Unknown content rejects the entire scene. The only admitted style is the exact inert root
 * marker "isolation:isolate"; it is not retained. No caller-owned object is mutated or frozen.
 * Revalidate data after crossing an untrusted transport; a TypeScript type is not validation.
 */
export function sanitizeVisioForeignVectorTree(
	root: unknown,
	options: Partial<VisioForeignVectorLimits> = {},
): VisioForeignVector {
	try {
		return sanitize(root, limitsFor(options));
	} catch (error) {
		if (error instanceof VisioForeignVectorError) throw error;
		// Do not expose converter content or arbitrary property-trap exception messages.
		unsafe('Unable to inspect converter output safely.');
	}
}
function sanitize(root: unknown, limits: VisioForeignVectorLimits): VisioForeignVector {
	let nodes = 0,
		characters = 0;
	const seen = new WeakSet<object>(),
		pathBudget = { operands: 0, commands: 0 };
	const clips: VisioForeignVectorClip[] = [],
		ids = new Map<string, number>();
	const pending: { target: ClipTarget; id: string }[] = [];
	const freezeLater: object[] = [];
	function claim(value: unknown, depth: number): void {
		if (!value || typeof value !== 'object' || seen.has(value))
			unsafe('Repeated or cyclic tree object.');
		if (depth > limits.maxDepth) limit('Tree depth limit exceeded.');
		if (++nodes > limits.maxNodes) limit('Tree node limit exceeded.');
		seen.add(value);
	}
	function charge(value: string): void {
		characters += value.length;
		if (characters > limits.maxCharacters) limit('Tree character limit exceeded.');
	}
	function node(
		value: unknown,
		depth: number,
		clipping = false,
	): { tag: string; attrs: RecordValue; children: unknown } {
		claim(value, depth);
		const data = record(value, ['tag', 'attrs', 'children']);
		if (typeof data.tag !== 'string' || !Object.hasOwn(ATTRS, data.tag))
			unsafe('Unsupported vector element.');
		charge(data.tag);
		const allowed =
			clipping && data.tag === 'path' ? ['d', 'clip-rule', ...BASE] : ATTRS[data.tag]!;
		const attrs = record(data.attrs, allowed);
		for (const key of Object.keys(attrs)) {
			charge(key);
			if (typeof attrs[key] === 'string') charge(attrs[key]);
			else if (typeof attrs[key] !== 'number') unsafe('Invalid attribute value type.');
		}
		return { tag: data.tag, attrs, children: data.children };
	}
	function children(value: unknown, optional = false): unknown[] {
		if (value === undefined && optional) return [];
		if (!Array.isArray(value) || Object.getPrototypeOf(value) !== Array.prototype)
			unsafe('Expected plain child array.');
		// Check length before keys, copies, mapping or sparse-array allocation.
		if (value.length > limits.maxNodes - nodes) limit('Child array node limit exceeded.');
		if (seen.has(value)) unsafe('Repeated child array.');
		seen.add(value);
		const keys = Reflect.ownKeys(value);
		if (keys.length !== value.length + 1) unsafe('Sparse or extended child array.');
		const output: unknown[] = [];
		for (let i = 0; i < value.length; i++) {
			const descriptor = Object.getOwnPropertyDescriptor(value, i);
			if (!descriptor || !('value' in descriptor) || !descriptor.enumerable)
				unsafe('Sparse or accessor child array.');
			output.push(descriptor.value as unknown);
		}
		return output;
	}
	function base(attrs: RecordValue): { matrix: ReturnType<typeof matrix>; clipIndex?: number } {
		const target: { matrix: ReturnType<typeof matrix>; clipIndex?: number } = {
			matrix: matrix(attrs.transform, limits.maxCoordinate),
		};
		if (attrs['clip-path'] !== undefined) {
			const value = attrs['clip-path'];
			const match =
				typeof value === 'string' ? /^url\(#([A-Za-z_][A-Za-z0-9_.-]{0,127})\)$/.exec(value) : null;
			if (!match) unsafe('Unsupported clip reference.');
			pending.push({ target, id: match[1]! });
		}
		return target;
	}
	function leaf(data: ReturnType<typeof node>): readonly ReturnType<typeof pathCommands>[number][] {
		if (children(data.children, true).length !== 0) unsafe('Path cannot contain child elements.');
		return pathCommands(data.attrs.d, limits, pathBudget);
	}
	function item(value: unknown, depth: number): VisioForeignVectorNode {
		const data = node(value, depth);
		if (data.tag !== 'g' && data.tag !== 'path') unsafe('Only paths and groups may be drawn.');
		const target = base(data.attrs);
		// Mutate only new output objects while resolving references, never converter-owned objects.
		const result: VisioForeignVectorNode =
			data.tag === 'path'
				? (Object.assign(target, {
						kind: 'path' as const,
						commands: leaf(data),
						paint: paint(data.attrs),
					}) satisfies VisioForeignVectorPath)
				: Object.assign(target, {
						kind: 'group' as const,
						items: Object.freeze(children(data.children).map((child) => item(child, depth + 1))),
					});
		freezeLater.push(result);
		return result;
	}
	function definition(value: unknown, depth: number): void {
		const data = node(value, depth);
		if (data.tag !== 'clipPath') unsafe('Only clip paths may be defined.');
		const id = data.attrs.id;
		if (typeof id !== 'string' || !ID.test(id) || ids.has(id))
			unsafe('Invalid or duplicate clip identity.');
		if (data.attrs.clipPathUnits !== undefined && data.attrs.clipPathUnits !== 'userSpaceOnUse')
			unsafe('Unsupported clip coordinate system.');
		const inheritedRule = rule(data.attrs['clip-rule']);
		ids.set(id, clips.length);
		const target = base(data.attrs);
		const paths: VisioForeignVectorClipPath[] = [];
		for (const child of children(data.children)) {
			const part = node(child, depth + 1, true);
			if (part.tag !== 'path') unsafe('Clip content must be a path.');
			const path = Object.assign(base(part.attrs), {
				commands: leaf(part),
				clipRule:
					part.attrs['clip-rule'] === undefined ? inheritedRule : rule(part.attrs['clip-rule']),
			});
			paths.push(path);
			freezeLater.push(path);
		}
		const clip = Object.assign(target, { items: Object.freeze(paths) });
		clips.push(clip);
		freezeLater.push(clip);
	}
	const data = node(root, 0);
	if (data.tag !== 'svg' || data.attrs.xmlns !== 'http://www.w3.org/2000/svg')
		unsafe('Invalid vector root.');
	if (data.attrs.style !== undefined && data.attrs.style !== 'isolation:isolate')
		unsafe('Unsupported root style.');
	const width = finite(data.attrs.width, limits.maxDimension),
		height = finite(data.attrs.height, limits.maxDimension);
	if (width < 1 || height < 1) unsafe('Empty vector dimensions.');
	const viewBox = numberList(data.attrs.viewBox, 4, limits.maxDimension);
	if (viewBox[0] !== 0 || viewBox[1] !== 0 || viewBox[2] !== width || viewBox[3] !== height)
		unsafe('Viewport differs from dimensions.');
	const items: VisioForeignVectorNode[] = [];
	for (const child of children(data.children)) {
		// Inspect the discriminator without invoking getters; claim and charge in the selected branch.
		const peek = record(child, ['tag', 'attrs', 'children']);
		if (peek.tag === 'defs') {
			const defs = node(child, 1);
			for (const value of children(defs.children)) definition(value, 2);
		} else items.push(item(child, 1));
	}
	for (const { target, id } of pending) {
		const index = ids.get(id);
		if (index === undefined) unsafe('Unresolved clip reference.');
		target.clipIndex = index;
	}
	const scene: VisioForeignVector = {
		kind: 'vector',
		width,
		height,
		items: Object.freeze(items),
		clips: Object.freeze(clips),
	};
	checkVectorGraph(scene, limits);
	for (const value of freezeLater) Object.freeze(value);
	return Object.freeze(scene);
}

export { validateVisioForeignVector } from './foreign-vector-validation.js';
