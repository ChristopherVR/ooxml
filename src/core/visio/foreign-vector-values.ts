import {
	VISIO_FOREIGN_VECTOR_LIMITS,
	limit,
	unsafe,
	type VisioForeignVectorLimits,
	type VisioForeignVectorMatrix,
	type VisioForeignVectorPaint,
	type VisioForeignVectorRule,
} from './foreign-vector-types.js';

export const IDENTITY: VisioForeignVectorMatrix = Object.freeze([1, 0, 0, 1, 0, 0]);
const NUMBER = /^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;
const LIST =
	/^[\x20\t\r\n]*[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?(?:[\x20\t\r\n]*,[\x20\t\r\n]*|[\x20\t\r\n]+)/;
export type RecordValue = Record<string, unknown>;
/** Reject accessors/exotic prototypes rather than invoking document-owned getters or coercions. */
export function record(value: unknown, allowed: readonly string[]): RecordValue {
	if (!value || typeof value !== 'object' || Array.isArray(value))
		unsafe('Expected a plain object.');
	const prototype: unknown = Object.getPrototypeOf(value);
	if (prototype !== Object.prototype && prototype !== null) unsafe('Unexpected object prototype.');
	const result: RecordValue = Object.create(null) as RecordValue;
	const keys = Reflect.ownKeys(value);
	if (keys.length > allowed.length) unsafe('Unexpected object fields.');
	for (const key of keys) {
		if (typeof key !== 'string' || !allowed.includes(key)) unsafe('Unexpected object field.');
		const descriptor = Object.getOwnPropertyDescriptor(value, key);
		if (!descriptor || !('value' in descriptor) || !descriptor.enumerable)
			unsafe('Invalid data property.');
		result[key] = descriptor.value as unknown;
	}
	return result;
}
export function limitsFor(options: Partial<VisioForeignVectorLimits>): VisioForeignVectorLimits {
	const result = { ...VISIO_FOREIGN_VECTOR_LIMITS };
	const values = record(options, Object.keys(result));
	for (const key of Object.keys(values) as (keyof VisioForeignVectorLimits)[]) {
		const value = values[key];
		if (
			typeof value !== 'number' ||
			!Number.isSafeInteger(value) ||
			value < 1 ||
			value > result[key]
		)
			unsafe('Limits must be positive integers no greater than the defaults.');
		result[key] = value;
	}
	return result;
}
export function finite(value: unknown, maximum: number): number {
	if (typeof value !== 'number' || !Number.isFinite(value)) unsafe('Expected a finite number.');
	if (Math.abs(value) > maximum) limit('Numeric magnitude limit exceeded.');
	return value;
}
export function scalar(value: unknown, maximum: number, minimum = 0): number {
	if (typeof value === 'string') {
		if (!NUMBER.test(value)) unsafe('Invalid numeric attribute.');
		value = Number(value);
	}
	const n = finite(value, maximum);
	if (n < minimum) unsafe('Numeric attribute below its minimum.');
	return n;
}
export function numberList(value: unknown, length: number, maximum: number): number[] {
	if (typeof value !== 'string') unsafe('Expected a numeric list.');
	const result: number[] = [];
	let rest = value.trim();
	for (let i = 0; i < length - 1; i++) {
		const match = LIST.exec(rest);
		if (!match) unsafe('Invalid numeric list.');
		const token = match[0].trim().replace(/,$/, '').trim();
		result.push(finite(Number(token), maximum));
		rest = rest.slice(match[0].length);
	}
	if (!NUMBER.test(rest)) unsafe('Invalid numeric list arity.');
	result.push(finite(Number(rest), maximum));
	return result;
}
export function matrix(value: unknown, maximum: number): VisioForeignVectorMatrix {
	if (value === undefined) return IDENTITY;
	if (typeof value !== 'string' || !/^matrix\([^()]*\)$/.test(value))
		unsafe('Unsupported affine transform.');
	return Object.freeze(
		numberList(value.slice(7, -1), 6, maximum),
	) as unknown as VisioForeignVectorMatrix;
}
export function compose(
	a: VisioForeignVectorMatrix,
	b: VisioForeignVectorMatrix,
	max: number,
): VisioForeignVectorMatrix {
	const result: VisioForeignVectorMatrix = [
		a[0] * b[0] + a[2] * b[1],
		a[1] * b[0] + a[3] * b[1],
		a[0] * b[2] + a[2] * b[3],
		a[1] * b[2] + a[3] * b[3],
		a[0] * b[4] + a[2] * b[5] + a[4],
		a[1] * b[4] + a[3] * b[5] + a[5],
	];
	for (const value of result) finite(value, max);
	return result;
}
export function rule(value: unknown): VisioForeignVectorRule {
	if (value === undefined) return 'nonzero';
	if (value !== 'nonzero' && value !== 'evenodd') unsafe('Unsupported winding rule.');
	return value;
}
function literal(value: unknown, fallback: string): string {
	if (value === undefined) return fallback;
	if (typeof value !== 'string' || !/^(?:none|#[0-9a-fA-F]{6})$/.test(value))
		unsafe('Nonliteral paint.');
	return value.toLowerCase();
}
function choice<T extends string>(value: unknown, choices: readonly T[], fallback: T): T {
	if (value === undefined) return fallback;
	if (typeof value !== 'string' || !choices.includes(value as T))
		unsafe('Unsupported paint value.');
	return value as T;
}
export function paint(attrs: RecordValue): VisioForeignVectorPaint {
	return Object.freeze({
		fill: literal(attrs.fill, '#000000'),
		stroke: literal(attrs.stroke, 'none'),
		fillRule: rule(attrs['fill-rule']),
		strokeWidth: scalar(attrs['stroke-width'] === undefined ? 1 : attrs['stroke-width'], 1024),
		strokeMiterlimit: scalar(
			attrs['stroke-miterlimit'] === undefined ? 4 : attrs['stroke-miterlimit'],
			1024,
			1,
		),
		strokeLinecap: choice(attrs['stroke-linecap'], ['butt', 'round', 'square'], 'butt'),
		strokeLinejoin: choice(attrs['stroke-linejoin'], ['miter', 'round', 'bevel'], 'miter'),
		opacity: scalar(attrs.opacity === undefined ? 1 : attrs.opacity, 1),
		fillOpacity: scalar(attrs['fill-opacity'] === undefined ? 1 : attrs['fill-opacity'], 1),
		strokeOpacity: scalar(attrs['stroke-opacity'] === undefined ? 1 : attrs['stroke-opacity'], 1),
	});
}
