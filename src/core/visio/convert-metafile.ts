import { inspectVisioEmfAdmission } from './emf-admission';
import { emfInputView } from './emf-admission-input';
import {
	VISIO_EMF_CONVERSION_LIMITS,
	VISIO_EMF_CONVERSION_MAX_DIMENSION,
} from './emf-admission-types';
import {
	sanitizeVisioForeignVectorTree,
	VisioForeignVectorError,
	type VisioForeignVector,
} from './foreign-vector';

/** Trusted package function, normally emf-converter's browser convertMetafileToSvgTree.
 * The host supplies the package so this format boundary cannot accidentally load Node codecs.
 * Invoke only inside a disposable worker with a parent-owned deadline and cancellation.
 */
export type VisioMetafileTreeConverter = (
	bytes: ArrayBuffer,
	options: {
		dpiScale: number;
		maxWidth: number;
		maxHeight: number;
		maxCanvasDimension: number;
		exactRasterOps: false;
		includeSize: true;
		idPrefix: string;
	},
) => Promise<unknown>;
export type VisioMetafileConversionResult =
	| { status: 'ok'; vector: VisioForeignVector }
	| {
			status: 'invalid' | 'unsupported' | 'budget-exceeded' | 'conversion-failed';
			code: string;
			message: string;
	  };
/** Package adapter, not a second EMF renderer. Never returns raw SVG or partial output.
 * Document import must explicitly inject the trusted converter in an isolated worker.
 * Post-conversion validation bounds retained output, not peak converter heap allocation.
 */
export async function convertVisioMetafile(
	input: Uint8Array,
	converter: VisioMetafileTreeConverter,
): Promise<VisioMetafileConversionResult> {
	const source = emfInputView(input);
	if (!source)
		return { status: 'invalid', code: 'input', message: 'Expected fixed, unshared EMF bytes.' };
	if (source.byteLength > VISIO_EMF_CONVERSION_LIMITS.maxInputBytes)
		return {
			status: 'budget-exceeded',
			code: 'input-limit',
			message: 'Conversion input exceeds 256 KiB.',
		};
	// Intrinsic view access and a private copy avoid caller iterators, accessors and mutation
	// between inspection and the converter's first asynchronous continuation.
	const bytes = new Uint8Array(source.byteLength);
	bytes.set(new Uint8Array(source.buffer, source.byteOffset, source.byteLength));
	const admission = inspectVisioEmfAdmission(bytes, VISIO_EMF_CONVERSION_LIMITS);
	if (admission.status !== 'admitted') {
		const note = admission.diagnostics[0];
		return {
			status: admission.status,
			code: note?.code ?? 'admission',
			message: note?.message ?? 'Metafile did not pass bounded preflight.',
		};
	}
	const header = admission.header!;
	const maxDimension = VISIO_EMF_CONVERSION_MAX_DIMENSION;
	if (header.pixelWidth > maxDimension || header.pixelHeight > maxDimension)
		return {
			status: 'budget-exceeded',
			code: 'dimension-limit',
			message: 'Conversion dimensions exceed 2048 pixels.',
		};
	try {
		const tree = await converter(bytes.buffer, {
			dpiScale: 1,
			maxWidth: maxDimension,
			maxHeight: maxDimension,
			maxCanvasDimension: maxDimension,
			exactRasterOps: false,
			includeSize: true,
			idPrefix: 'visio-emf-',
		});
		if (tree === null)
			return {
				status: 'conversion-failed',
				code: 'empty-output',
				message: 'Converter returned no vector output.',
			};
		const vector = sanitizeVisioForeignVectorTree(tree, {
			maxNodes: 2048,
			maxPathOperands: 20000,
			maxPathCommands: 10000,
			maxCharacters: 256 * 1024,
			maxExpandedNodes: 2048,
			maxExpandedOperands: 20000,
			maxExpandedCommands: 10000,
		});
		if (vector.width !== header.pixelWidth || vector.height !== header.pixelHeight)
			return {
				status: 'conversion-failed',
				code: 'dimension-mismatch',
				message: 'Converter output differs from inclusive EMF bounds.',
			};
		return { status: 'ok', vector };
	} catch (error) {
		return {
			status:
				error instanceof VisioForeignVectorError && error.code === 'vector-limit'
					? 'budget-exceeded'
					: 'conversion-failed',
			code: error instanceof VisioForeignVectorError ? error.code : 'converter-error',
			message: 'Metafile conversion failed or produced inadmissible vector output.',
		};
	}
}
