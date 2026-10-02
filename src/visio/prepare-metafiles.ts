import { convertVisioMetafile, type VisioMetafileTreeConverter } from './convert-metafile.js';
import type { VisioForeignVector } from './foreign-vector.js';
import { NS } from '../xml/index.js';
import { inspectVisioEmfAdmission } from './emf-admission.js';
import type { VisioPackage } from './package.js';
import { children, type Report } from './sheet.js';

/** Inspection budgets across unique internal media parts, independently of ZIP/XML limits. */
export const VISIO_METAFILE_INSPECTION_LIMITS = Object.freeze({
	maxAssets: 32,
	maxTotalBytes: 32 * 1024 * 1024,
	maxTotalRecords: 100_000,
});
export const VISIO_METAFILE_CONVERSION_LIMITS = Object.freeze({
	maxAssets: 8,
	maxTotalBytes: 1024 * 1024,
});
type Prepared = { notes: readonly Note[]; vector?: VisioForeignVector };
type Note = { code: string; message: string };
interface State {
	bytes: number;
	records: number;
	results: Map<string, Promise<Prepared>>;
	conversionAssets: number;
	conversionBytes: number;
	tail: Promise<void>;
}
const states = new WeakMap<VisioPackage, State>();
/** Conversion is opt-in and must be hosted in a disposable worker with a parent-owned deadline. */
export async function inspectEmbeddedVisioMetafile(
	pkg: VisioPackage,
	sourcePart: string,
	foreignData: Element,
	report: Report,
	converter?: VisioMetafileTreeConverter,
): Promise<VisioForeignVector | undefined> {
	const refs = children(foreignData, 'Rel');
	const id = refs.length === 1 ? refs[0]?.getAttributeNS(NS.r, 'id') : undefined;
	const rel = id ? (await pkg.relationships(sourcePart)).get(id) : undefined;
	if (!rel) {
		report('invalid-image-relationship', 'Metafile has no unambiguous relationship reference.');
		return;
	}
	if (rel.mode !== 'Internal') {
		report('external-image', 'External metafile relationships are not loaded.');
		return;
	}
	if (rel.type !== `${NS.r}/image`) {
		report('invalid-image-relationship', 'Metafile does not reference an image relationship.');
		return;
	}
	let state = states.get(pkg);
	if (!state) {
		state = {
			bytes: 0,
			records: 0,
			conversionAssets: 0,
			conversionBytes: 0,
			results: new Map(),
			tail: Promise.resolve(),
		};
		states.set(pkg, state);
	}
	let pending = state.results.get(rel.target);
	if (!pending) {
		if (state.results.size >= VISIO_METAFILE_INSPECTION_LIMITS.maxAssets) {
			report(
				'emf-document-limit',
				'Additional metafile parts were not inspected because the document asset limit was reached.',
				{ part: rel.target },
			);
			return;
		}
		const current = state;
		pending = current.tail.then(() => inspectPart(pkg, rel.target, current, converter));
		// Reserve a unique slot before yielding; duplicate calls share the same promise.
		current.results.set(rel.target, pending);
		current.tail = pending.then(
			() => undefined,
			() => undefined,
		);
	}
	const prepared = await pending;
	for (const note of prepared.notes) report(note.code, note.message, { part: rel.target });
	return prepared.vector;
}

async function inspectPart(
	pkg: VisioPackage,
	part: string,
	state: State,
	converter?: VisioMetafileTreeConverter,
): Promise<Prepared> {
	let notes: readonly Note[];
	const size = pkg.getPartByteLength(part);
	const remaining = VISIO_METAFILE_INSPECTION_LIMITS.maxTotalRecords - state.records;
	if (size > 4 * 1024 * 1024) {
		notes = [
			{
				code: 'emf-input-limit',
				message: 'Metafile bytes exceed the per-asset limit; the part was not inflated.',
			},
		];
	} else if (
		size > VISIO_METAFILE_INSPECTION_LIMITS.maxTotalBytes - state.bytes ||
		remaining <= 0
	) {
		notes = [
			{
				code: 'emf-document-limit',
				message:
					'Metafile inspection exceeded the aggregate byte or record limit; the part was not inflated.',
			},
		];
	} else {
		const bytes = await pkg.readBytes(part);
		state.bytes += bytes.byteLength;
		const result = inspectVisioEmfAdmission(bytes, { maxRecords: Math.min(20_000, remaining) });
		state.records += result.metrics.records;
		notes = result.diagnostics.map((note) => ({
			code: `emf-${note.code}`,
			message: `${note.message} Record ${note.recordType ?? 'header'} at byte ${note.offset}.`,
		}));
		if (result.omittedDiagnostics)
			notes = [
				...notes,
				{
					code: 'emf-diagnostics-truncated',
					message: `${result.omittedDiagnostics} additional metafile inspection diagnostics were omitted.`,
				},
			];
		if (result.status === 'admitted' && converter) {
			if (
				state.conversionAssets >= VISIO_METAFILE_CONVERSION_LIMITS.maxAssets ||
				bytes.byteLength > VISIO_METAFILE_CONVERSION_LIMITS.maxTotalBytes - state.conversionBytes
			) {
				return {
					notes: [
						{
							code: 'emf-conversion-document-limit',
							message: 'Metafile conversion exceeds the document asset or byte budget.',
						},
					],
				};
			}
			state.conversionAssets++;
			state.conversionBytes += bytes.byteLength;
			const converted = await convertVisioMetafile(bytes, converter);
			if (converted.status === 'ok')
				return {
					notes: [
						{
							code: 'emf-limited-rendering',
							message:
								'Rendered the bounded EMF primitive subset; Microsoft Visio fidelity is not established.',
						},
					],
					vector: converted.vector,
				};
			return { notes: [{ code: `emf-${converted.code}`, message: converted.message }] };
		}
		if (result.status === 'admitted')
			notes = [
				...notes,
				{
					code: 'emf-rendering-disabled',
					message:
						'This metafile passes the narrow structural preflight, but vector conversion and fidelity are not enabled.',
				},
			];
	}
	return { notes };
}
