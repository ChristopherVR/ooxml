import {
	VISIO_EMF_ADMISSION_LIMITS,
	type VisioEmfAdmissionOptions,
} from './emf-admission-types.js';

// Capture the actual TypedArray accessors. Own accessors, proxies and toStringTag overrides
// must never control buffer boundaries or execute while byte records are inspected.
const typedArray = Object.getPrototypeOf(Uint8Array.prototype) as object;
const bufferGetter = Object.getOwnPropertyDescriptor(typedArray, 'buffer')!.get!;
const offsetGetter = Object.getOwnPropertyDescriptor(typedArray, 'byteOffset')!.get!;
const lengthGetter = Object.getOwnPropertyDescriptor(typedArray, 'byteLength')!.get!;
const tagGetter = Object.getOwnPropertyDescriptor(typedArray, Symbol.toStringTag)!.get!;
const resizableGetter = Object.getOwnPropertyDescriptor(ArrayBuffer.prototype, 'resizable')?.get;

export function emfInputView(input: unknown): DataView<ArrayBuffer> | null {
	try {
		if (tagGetter.call(input) !== 'Uint8Array') return null;
		const buffer: unknown = bufferGetter.call(input);
		if (!(buffer instanceof ArrayBuffer) || resizableGetter?.call(buffer) === true) return null;
		return new DataView(
			buffer,
			offsetGetter.call(input) as number,
			lengthGetter.call(input) as number,
		);
	} catch {
		return null;
	}
}

/** Read only known own data properties, before inspecting the source byte buffer.
 * Accessors/inherited configuration are rejected. Proxy traps can run as part of descriptor
 * lookup, but their exceptions cannot escape and they cannot detach an already-inspected view.
 */
export function emfInputOptions(options: unknown): VisioEmfAdmissionOptions | null {
	if (!options || typeof options !== 'object') return null;
	try {
		const prototype: unknown = Object.getPrototypeOf(options);
		if (prototype !== null && prototype !== Object.prototype) return null;
		const snapshot: { -readonly [K in keyof VisioEmfAdmissionOptions]: number } = {};
		for (const key of Object.keys(
			VISIO_EMF_ADMISSION_LIMITS,
		) as (keyof VisioEmfAdmissionOptions)[]) {
			const descriptor = Object.getOwnPropertyDescriptor(options, key);
			if (!descriptor) continue;
			if (!('value' in descriptor)) return null;
			const value: unknown = descriptor.value;
			if (value === undefined) continue;
			if (typeof value !== 'number') return null;
			snapshot[key] = value;
		}
		return snapshot;
	} catch {
		return null;
	}
}
